#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
AIword - 智能粘贴工具
将 AI 生成的内容（Markdown 表格、HTML 富文本）智能粘贴到 Word/WPS/Excel
"""

import sys
import os
import threading
import pythoncom

from config_manager import ConfigManager
from logger import Logger
from hotkey_manager import HotkeyManager
from content_processor import ContentProcessor
from tray_icon import TrayIcon


class AIwordApp:
    def __init__(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_dir, 'config.json')
        
        self.config_manager = ConfigManager(config_path)
        
        log_dir = self.config_manager.get('log_dir', './logs')
        if not os.path.isabs(log_dir):
            log_dir = os.path.join(base_dir, log_dir)
        
        self.logger = Logger(log_dir)
        self.logger.info("=" * 50)
        self.logger.info(f"AIword 启动中... 版本: {self.config_manager.get('version', '1.0.0')}")
        
        self.content_processor = ContentProcessor(self.config_manager, self.logger)
        
        default_hotkey = self.config_manager.get('hotkey', 'ctrl+alt+v')
        self.hotkey_manager = HotkeyManager(self.logger, default_hotkey)
        self.hotkey_manager.register_callback(self._on_hotkey_pressed)
        
        if not self.config_manager.get('hotkey_enabled', True):
            self.hotkey_manager.disable()
        
        self.tray_icon = TrayIcon(
            config_manager=self.config_manager,
            logger=self.logger,
            hotkey_manager=self.hotkey_manager,
            on_exit=self._on_exit,
            on_process=self._process_content
        )
        
        self.running = True
    
    def _on_hotkey_pressed(self):
        self.logger.info("检测到热键按下")
        threading.Thread(target=self._process_content, daemon=True).start()
    
    def _process_content(self):
        try:
            pythoncom.CoInitialize()
            self.content_processor.process_clipboard()
        except Exception as e:
            self.logger.error(f"处理内容时出错: {str(e)}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _on_exit(self):
        self.logger.info("正在停止 AIword...")
        self.running = False
        self.hotkey_manager.disable()
        self.logger.info("AIword 已退出")
        sys.exit(0)
    
    def run(self):
        self.logger.info("AIword 已启动，正在运行中...")
        self.logger.info(f"当前热键: {self.hotkey_manager.get_current_hotkey()}")
        self.logger.info("提示: 点击系统托盘图标进行操作")
        
        self.tray_icon.run()


def main():
    try:
        app = AIwordApp()
        app.run()
    except Exception as e:
        print(f"程序启动失败: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
