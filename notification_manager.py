import ctypes
from typing import Optional
from logger import Logger


class NotificationManager:
    def __init__(self, logger: Logger, enabled: bool = True):
        self.logger = logger
        self.enabled = enabled
    
    def set_enabled(self, enabled: bool):
        self.enabled = enabled
    
    def show(self, title: str, message: str, icon_type: str = "info") -> bool:
        if not self.enabled:
            return False
        
        try:
            icon_map = {
                "info": 0x40,
                "warning": 0x30,
                "error": 0x10
            }
            
            icon_flags = icon_map.get(icon_type, 0x40)
            
            messagebox = ctypes.windll.user32.MessageBoxW
            result = messagebox(None, message, title, icon_flags | 0x40000)
            
            self.logger.info(f"已显示通知: {title} - {message}")
            return True
            
        except Exception as e:
            self.logger.error(f"显示通知失败: {str(e)}")
            return False
    
    def show_success(self, title: str = "操作成功", message: str = "") -> bool:
        return self.show(title, message, "info")
    
    def show_warning(self, title: str = "警告", message: str = "") -> bool:
        return self.show(title, message, "warning")
    
    def show_error(self, title: str = "错误", message: str = "") -> bool:
        return self.show(title, message, "error")
