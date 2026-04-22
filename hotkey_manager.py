import keyboard
from typing import Callable, Optional
from logger import Logger


class HotkeyManager:
    def __init__(self, logger: Logger, default_hotkey: str = "ctrl+alt+v"):
        self.logger = logger
        self.current_hotkey = default_hotkey
        self.enabled = True
        self.callback = None
        self.hotkey_id = None
    
    def set_hotkey(self, hotkey: str) -> bool:
        try:
            if self.hotkey_id is not None:
                keyboard.remove_hotkey(self.hotkey_id)
            
            self.current_hotkey = hotkey
            
            if self.callback and self.enabled:
                self.hotkey_id = keyboard.add_hotkey(hotkey, self.callback)
            
            self.logger.info(f"热键已设置为: {hotkey}")
            return True
        except Exception as e:
            self.logger.error(f"设置热键失败: {str(e)}")
            return False
    
    def register_callback(self, callback: Callable) -> None:
        self.callback = callback
        if self.enabled:
            try:
                if self.hotkey_id is not None:
                    keyboard.remove_hotkey(self.hotkey_id)
                self.hotkey_id = keyboard.add_hotkey(self.current_hotkey, callback)
                self.logger.info(f"热键回调已注册: {self.current_hotkey}")
            except Exception as e:
                self.logger.error(f"注册热键回调失败: {str(e)}")
    
    def enable(self) -> None:
        if not self.enabled:
            self.enabled = True
            if self.callback:
                try:
                    self.hotkey_id = keyboard.add_hotkey(self.current_hotkey, self.callback)
                    self.logger.info("热键已启用")
                except Exception as e:
                    self.logger.error(f"启用热键失败: {str(e)}")
    
    def disable(self) -> None:
        if self.enabled:
            self.enabled = False
            if self.hotkey_id is not None:
                try:
                    keyboard.remove_hotkey(self.hotkey_id)
                    self.hotkey_id = None
                    self.logger.info("热键已禁用")
                except Exception as e:
                    self.logger.error(f"禁用热键失败: {str(e)}")
    
    def get_current_hotkey(self) -> str:
        return self.current_hotkey
    
    def is_enabled(self) -> bool:
        return self.enabled
