import os
import tempfile
import pythoncom
import win32com.client
from typing import Optional, Tuple
from logger import Logger


class WordIntegration:
    def __init__(self, logger: Logger):
        self.logger = logger
        self.word_app = None
        self.wps_app = None
    
    def _get_running_word_app(self):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.GetActiveObject("Word.Application")
            self.logger.info("已连接到已运行的 Word 实例")
            return app
        except Exception as e:
            self.logger.debug(f"没有找到已运行的 Word 实例: {str(e)}")
            return None
    
    def _get_running_wps_app(self):
        try:
            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("Kwps.Application")
                self.logger.info("已连接到已运行的 WPS 实例")
                return app
            except:
                app = win32com.client.GetActiveObject("Wps.Application")
                self.logger.info("已连接到已运行的 WPS 实例")
                return app
        except Exception as e:
            self.logger.debug(f"没有找到已运行的 WPS 实例: {str(e)}")
            return None
    
    def _start_word_app(self):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = True
            self.logger.info("已启动新的 Word 实例")
            return app
        except Exception as e:
            self.logger.error(f"无法启动 Word: {str(e)}")
            return None
    
    def _start_wps_app(self):
        try:
            pythoncom.CoInitialize()
            try:
                app = win32com.client.Dispatch("Kwps.Application")
                app.Visible = True
                self.logger.info("已启动新的 WPS 实例")
                return app
            except:
                app = win32com.client.Dispatch("Wps.Application")
                app.Visible = True
                self.logger.info("已启动新的 WPS 实例")
                return app
        except Exception as e:
            self.logger.error(f"无法启动 WPS: {str(e)}")
            return None
    
    def get_running_app(self):
        word_app = self._get_running_word_app()
        if word_app:
            return word_app
        wps_app = self._get_running_wps_app()
        if wps_app:
            return wps_app
        return None
    
    def start_app(self, prefer_wps: bool = True):
        if prefer_wps:
            wps_app = self._start_wps_app()
            if wps_app:
                return wps_app
            word_app = self._start_word_app()
            return word_app
        else:
            word_app = self._start_word_app()
            if word_app:
                return word_app
            wps_app = self._start_wps_app()
            return wps_app
    
    def get_active_document(self, allow_start: bool = False):
        app = self.get_running_app()
        
        if not app and allow_start:
            app = self.start_app()
        
        if app:
            try:
                return app.ActiveDocument
            except Exception as e:
                self.logger.debug(f"获取活动文档失败: {str(e)}")
                pass
        
        return None
    
    def is_available(self) -> bool:
        return self.get_running_app() is not None
    
    def insert_html(self, html_content: str, move_cursor_to_end: bool = True) -> bool:
        try:
            doc = self.get_active_document()
            if not doc:
                self.logger.warning("没有活动的文档")
                return False
            
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                html_full = f'''
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
</head>
<body>
{html_content}
</body>
</html>
'''
                f.write(html_full)
                temp_path = f.name
            
            try:
                selection = doc.Application.Selection
                
                if move_cursor_to_end:
                    selection.EndKey(Unit=6)
                
                temp_path_abs = os.path.abspath(temp_path)
                selection.InsertFile(temp_path_abs)
                
                self.logger.info("HTML 内容已成功插入到文档")
                return True
            finally:
                if os.path.exists(temp_path):
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
                    
        except Exception as e:
            self.logger.error(f"插入 HTML 内容失败: {str(e)}")
            return False
    
    def insert_text(self, text_content: str, move_cursor_to_end: bool = True) -> bool:
        try:
            doc = self.get_active_document()
            if not doc:
                self.logger.warning("没有活动的文档")
                return False
            
            selection = doc.Application.Selection
            
            if move_cursor_to_end:
                selection.EndKey(Unit=6)
            
            selection.TypeText(text_content)
            
            self.logger.info("文本内容已成功插入到文档")
            return True
            
        except Exception as e:
            self.logger.error(f"插入文本内容失败: {str(e)}")
            return False
    
    def save_as_docx(self, content: str, save_dir: str, html_format: bool = True) -> Optional[str]:
        try:
            if not os.path.exists(save_dir):
                os.makedirs(save_dir)
            
            word_app = self.get_running_app()
            if not word_app:
                word_app = self.start_app()
            
            if not word_app:
                self.logger.error("无法启动 Word 或 WPS")
                return None
            
            timestamp = self._get_timestamp()
            filename = f"AIword_{timestamp}.docx"
            filepath = os.path.join(save_dir, filename)
            
            doc = word_app.Documents.Add()
            
            if html_format:
                with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                    f.write(content)
                    temp_path = f.name
                
                try:
                    doc.Content.InsertFile(os.path.abspath(temp_path))
                finally:
                    if os.path.exists(temp_path):
                        try:
                            os.unlink(temp_path)
                        except:
                            pass
            else:
                doc.Content.Text = content
            
            doc.SaveAs2(os.path.abspath(filepath))
            doc.Close()
            
            self.logger.info(f"文档已保存到: {filepath}")
            return filepath
            
        except Exception as e:
            self.logger.error(f"保存文档失败: {str(e)}")
            return None
    
    def _get_timestamp(self) -> str:
        from datetime import datetime
        return datetime.now().strftime('%Y%m%d_%H%M%S')
