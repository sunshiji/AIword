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
import signal
import tkinter as tk

from config_manager import ConfigManager
from logger import Logger
from hotkey_manager import HotkeyManager
from content_processor import ContentProcessor
from tray_icon import TrayIcon
from dialog_manager import initialize_dialog_manager, get_dialog_manager


class AIwordApp:
    def __init__(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(self.base_dir, 'config.json')
        
        self.config_manager = ConfigManager(config_path)
        
        log_dir = self.config_manager.get('log_dir', './logs')
        if not os.path.isabs(log_dir):
            log_dir = os.path.join(self.base_dir, log_dir)
        
        self.logger = Logger(log_dir)
        self.logger.info("=" * 50)
        self.logger.info(f"AIword 启动中... 版本: {self.config_manager.get('version', '1.0.0')}")
        
        self.root = None
        self.tray_thread = None
        self.running = True
        self.processing_lock = threading.Lock()
        self.is_processing = False
        
        self._setup_signal_handlers()
        
        self.content_processor = ContentProcessor(self.config_manager, self.logger)
        
        default_hotkey = self.config_manager.get('hotkey', 'ctrl+alt+v')
        self.hotkey_manager = HotkeyManager(self.logger, default_hotkey)
        self.hotkey_manager.register_callback(self._on_hotkey_pressed)
        
        self.exit_hotkey_id = None
        self._register_exit_hotkey()
        
        if not self.config_manager.get('hotkey_enabled', True):
            self.hotkey_manager.disable()
        
        self.tray_icon = TrayIcon(
            config_manager=self.config_manager,
            logger=self.logger,
            hotkey_manager=self.hotkey_manager,
            on_exit=self._on_exit,
            on_process=self._queue_process_content
        )
    
    def _setup_signal_handlers(self):
        def signal_handler(sig, frame):
            self.logger.info(f"收到信号: {sig}")
            self._on_exit()
        
        try:
            signal.signal(signal.SIGINT, signal_handler)
            signal.signal(signal.SIGTERM, signal_handler)
        except:
            pass
    
    def _register_exit_hotkey(self):
        try:
            import keyboard
            self.exit_hotkey_id = keyboard.add_hotkey(
                'ctrl+alt+q',
                self._on_emergency_exit,
                suppress=False
            )
            self.logger.info("紧急退出热键已注册: Ctrl+Alt+Q")
        except Exception as e:
            self.logger.error(f"注册紧急退出热键失败: {str(e)}")
    
    def _unregister_exit_hotkey(self):
        if self.exit_hotkey_id:
            try:
                import keyboard
                keyboard.remove_hotkey(self.exit_hotkey_id)
                self.exit_hotkey_id = None
                self.logger.info("紧急退出热键已注销")
            except Exception as e:
                self.logger.error(f"注销紧急退出热键失败: {str(e)}")
    
    def _on_emergency_exit(self):
        self.logger.warning("紧急退出热键被触发")
        self._on_exit()
    
    def _on_hotkey_pressed(self):
        self.logger.info("检测到热键按下: Ctrl+Alt+V")
        self._queue_process_content()
    
    def _queue_process_content(self):
        with self.processing_lock:
            if self.is_processing:
                self.logger.warning("上一次处理尚未完成，跳过本次请求")
                return
            self.is_processing = True
        
        threading.Thread(target=self._process_content_wrapper, daemon=True).start()
    
    def _process_content_wrapper(self):
        try:
            self._process_content()
        finally:
            with self.processing_lock:
                self.is_processing = False
    
    def _process_content(self):
        try:
            pythoncom.CoInitialize()
            self.content_processor.process_clipboard()
        except Exception as e:
            self.logger.error(f"处理内容时出错: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _on_exit(self):
        if not self.running:
            return
        
        self.logger.info("正在停止 AIword...")
        self.running = False
        
        self.hotkey_manager.disable()
        self._unregister_exit_hotkey()
        
        try:
            dm = get_dialog_manager()
            dm.stop()
        except:
            pass
        
        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except:
                pass
        
        if self.root:
            try:
                self.root.quit()
                self.root.destroy()
            except:
                pass
        
        self.logger.info("AIword 已退出")
        os._exit(0)
    
    def _setup_tkinter(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title("AIword")
        
        initialize_dialog_manager(self.root, self.logger)
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_exit)
        
        self.logger.info("Tkinter 已初始化")
    
    def _run_tray_in_thread(self):
        def tray_runner():
            try:
                self.tray_icon.run()
            except Exception as e:
                self.logger.error(f"系统托盘错误: {str(e)}")
                import traceback
                self.logger.error(traceback.format_exc())
        
        self.tray_thread = threading.Thread(target=tray_runner, daemon=True)
        self.tray_thread.start()
        self.logger.info("系统托盘已在独立线程启动")
    
    def run(self):
        self.logger.info("AIword 已启动，正在运行中...")
        self.logger.info(f"当前热键: {self.hotkey_manager.get_current_hotkey()}")
        self.logger.info("紧急退出热键: Ctrl+Alt+Q")
        self.logger.info("提示: 点击系统托盘图标进行操作")
        
        self._setup_tkinter()
        self._run_tray_in_thread()
        
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.logger.info("收到键盘中断")
            self._on_exit()
        except Exception as e:
            self.logger.error(f"主循环错误: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            self._on_exit()


def main():
    try:
        app = AIwordApp()
        app.run()
    except Exception as e:
        print(f"程序启动失败: {str(e)}")
        import traceback
        traceback.print_exc()
        
        try:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'crash.log'), 'a') as f:
                f.write(f"\n{'='*50}\n")
                f.write(f"崩溃时间: {__import__('datetime').datetime.now()}\n")
                f.write(traceback.format_exc())
        except:
            pass
        
        sys.exit(1)


if __name__ == '__main__':
    main()
