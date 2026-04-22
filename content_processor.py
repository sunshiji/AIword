import os
import subprocess
from typing import Optional, Tuple
from enum import Enum

from config_manager import ConfigManager
from logger import Logger
from content_converter import ContentConverter
from clipboard_manager import ClipboardManager
from word_integration import WordIntegration
from excel_integration import ExcelIntegration
from notification_manager import NotificationManager
from confirm_dialog import ConfirmAction, show_confirm_dialog


class NoAppAction(Enum):
    AUTO_OPEN = "auto_open"
    SAVE_ONLY = "save_only"
    COPY_TO_CLIPBOARD = "copy_to_clipboard"
    NO_ACTION = "no_action"


class ContentProcessor:
    def __init__(self, config_manager: ConfigManager, logger: Logger):
        self.config_manager = config_manager
        self.logger = logger
        
        self.content_converter = ContentConverter(
            html_formatting=config_manager.get('html_formatting', True)
        )
        self.clipboard_manager = ClipboardManager(logger)
        self.word_integration = WordIntegration(logger)
        self.excel_integration = ExcelIntegration(logger)
        self.notification_manager = NotificationManager(
            logger,
            enabled=config_manager.get('notification_enabled', True)
        )
    
    def reload_config(self):
        self.config_manager.reload()
        self.content_converter.html_formatting = self.config_manager.get('html_formatting', True)
        self.notification_manager.set_enabled(self.config_manager.get('notification_enabled', True))
        self.logger.info("配置已重新加载")
    
    def process_clipboard(self) -> bool:
        content = self.clipboard_manager.get_text()
        
        if not content or not content.strip():
            self.logger.warning("剪贴板为空")
            self.notification_manager.show_warning("提示", "剪贴板为空")
            return False
        
        return self.process_content(content)
    
    def process_content(self, content: str) -> bool:
        result = self.content_converter.convert(content)
        content_type = result['type']
        
        self.logger.info(f"检测到内容类型: {content_type}")
        
        confirm_before_paste = self.config_manager.get('confirm_before_paste', True)
        
        if confirm_before_paste:
            word_available = self.word_integration.is_available()
            excel_available = self.excel_integration.is_available()
            no_app_action = self.config_manager.get('no_app_action', 'copy_to_clipboard')
            
            action = show_confirm_dialog(
                logger=self.logger,
                content=result['plain_text'] or content,
                content_type=content_type,
                html_content=result.get('html'),
                csv_content=result.get('csv'),
                word_available=word_available,
                excel_available=excel_available,
                no_app_action=no_app_action
            )
            
            return self._handle_confirm_action(action, result, content_type)
        else:
            if content_type == 'markdown_table':
                return self._process_table(result, allow_start=False)
            else:
                return self._process_text(result, allow_start=False)
    
    def _handle_confirm_action(self, action: ConfirmAction, result: dict, content_type: str) -> bool:
        if action == ConfirmAction.CANCEL:
            self.logger.info("用户取消操作")
            return True
        
        elif action == ConfirmAction.COPY_TO_CLIPBOARD:
            success = self.clipboard_manager.set_text(result['plain_text'])
            if success:
                self.notification_manager.show_success("操作完成", "内容已复制到剪贴板")
            return success
        
        elif action == ConfirmAction.SAVE_FILE:
            return self._save_to_file(
                content_type=content_type,
                content=result['plain_text'],
                html_content=result.get('html'),
                csv_content=result.get('csv')
            )
        
        elif action == ConfirmAction.CONFIRM:
            if content_type == 'markdown_table':
                return self._process_table(result, allow_start=True)
            else:
                return self._process_text(result, allow_start=True)
        
        return False
    
    def _process_table(self, result: dict, allow_start: bool = False) -> bool:
        csv_content = result['csv']
        html_content = result['html']
        
        if self.excel_integration.is_available():
            success = self.excel_integration.insert_csv(csv_content, allow_start=False)
            if success:
                self.notification_manager.show_success("操作成功", "Markdown 表格已粘贴到 Excel")
                self.logger.info("表格已成功插入到 Excel")
            return success
        elif allow_start:
            success = self.excel_integration.insert_csv(csv_content, allow_start=True)
            if success:
                self.notification_manager.show_success("操作成功", "已打开 Excel 并粘贴表格")
            return success
        else:
            return self._handle_no_app_action(
                content_type='table',
                content=csv_content,
                html_content=html_content,
                csv_content=csv_content
            )
    
    def _process_text(self, result: dict, allow_start: bool = False) -> bool:
        html_content = result['html']
        plain_text = result['plain_text']
        
        if self.word_integration.is_available():
            move_cursor = self.config_manager.get('move_cursor_to_end', True)
            
            if html_content and result['type'] != 'plain_text':
                success = self.word_integration.insert_html(html_content, move_cursor)
            else:
                success = self.word_integration.insert_text(plain_text, move_cursor)
            
            if success:
                self.notification_manager.show_success("操作成功", "内容已粘贴到 Word/WPS")
                self.logger.info("内容已成功插入到 Word/WPS")
            return success
        elif allow_start:
            move_cursor = self.config_manager.get('move_cursor_to_end', True)
            doc = self.word_integration.get_active_document(allow_start=True)
            
            if doc:
                if html_content and result['type'] != 'plain_text':
                    success = self.word_integration.insert_html(html_content, move_cursor)
                else:
                    success = self.word_integration.insert_text(plain_text, move_cursor)
                
                if success:
                    self.notification_manager.show_success("操作成功", "已打开 Word/WPS 并粘贴内容")
                return success
            else:
                return self._handle_no_app_action(
                    content_type='text',
                    content=plain_text,
                    html_content=html_content
                )
        else:
            return self._handle_no_app_action(
                content_type='text',
                content=plain_text,
                html_content=html_content
            )
    
    def _handle_no_app_action(self, content_type: str, content: str, 
                               html_content: str = None, csv_content: str = None) -> bool:
        action = self.config_manager.get('no_app_action', 'copy_to_clipboard')
        
        self.logger.info(f"未检测到目标应用，执行动作: {action}")
        
        if action == NoAppAction.AUTO_OPEN.value:
            return self._auto_open_app(content_type, content, html_content, csv_content)
        
        elif action == NoAppAction.SAVE_ONLY.value:
            return self._save_to_file(content_type, content, html_content, csv_content)
        
        elif action == NoAppAction.COPY_TO_CLIPBOARD.value:
            success = self.clipboard_manager.set_text(content)
            if success:
                self.notification_manager.show_success("操作完成", "内容已复制到剪贴板")
            return success
        
        elif action == NoAppAction.NO_ACTION.value:
            self.notification_manager.show_warning("提示", "未检测到目标应用，已取消操作")
            return True
        
        return False
    
    def _auto_open_app(self, content_type: str, content: str, 
                       html_content: str = None, csv_content: str = None) -> bool:
        if content_type == 'table':
            try:
                success = self.excel_integration.insert_csv(csv_content or content, allow_start=True)
                if success:
                    self.notification_manager.show_success("操作成功", "已打开 Excel 并粘贴表格")
                return success
            except Exception as e:
                self.logger.error(f"打开 Excel 失败: {str(e)}")
                return self._save_to_file(content_type, content, html_content, csv_content)
        else:
            try:
                doc = self.word_integration.get_active_document(allow_start=True)
                if doc:
                    move_cursor = self.config_manager.get('move_cursor_to_end', True)
                    if html_content:
                        success = self.word_integration.insert_html(html_content, move_cursor)
                    else:
                        success = self.word_integration.insert_text(content, move_cursor)
                    
                    if success:
                        self.notification_manager.show_success("操作成功", "已打开 Word/WPS 并粘贴内容")
                    return success
                else:
                    return self._save_to_file(content_type, content, html_content)
            except Exception as e:
                self.logger.error(f"打开 Word/WPS 失败: {str(e)}")
                return self._save_to_file(content_type, content, html_content)
    
    def _save_to_file(self, content_type: str, content: str, 
                       html_content: str = None, csv_content: str = None) -> bool:
        save_dir = self.config_manager.get('save_dir', './output')
        keep_files = self.config_manager.get('keep_generated_files', False)
        
        if not os.path.isabs(save_dir):
            save_dir = os.path.abspath(save_dir)
        
        try:
            if content_type == 'table':
                filepath = self.excel_integration.save_as_xlsx(csv_content or content, save_dir)
            else:
                filepath = self.word_integration.save_as_docx(
                    html_content or content, 
                    save_dir, 
                    html_format=bool(html_content)
                )
            
            if filepath:
                self.notification_manager.show_success("文件已保存", f"文件已保存到:\n{filepath}")
                return True
            else:
                return self.clipboard_manager.set_text(content)
                
        except Exception as e:
            self.logger.error(f"保存文件失败: {str(e)}")
            return self.clipboard_manager.set_text(content)
