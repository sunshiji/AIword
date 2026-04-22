import pyperclip
from typing import Optional
from logger import Logger


class ClipboardManager:
    def __init__(self, logger: Logger):
        self.logger = logger
        self._last_content = None
    
    def get_text(self) -> Optional[str]:
        try:
            content = pyperclip.paste()
            return content
        except Exception as e:
            self.logger.error(f"读取剪贴板失败: {str(e)}")
            return None
    
    def set_text(self, text: str) -> bool:
        try:
            pyperclip.copy(text)
            self.logger.info("内容已复制到剪贴板")
            return True
        except Exception as e:
            self.logger.error(f"写入剪贴板失败: {str(e)}")
            return False
    
    def has_changed(self) -> bool:
        current_content = self.get_text()
        if current_content != self._last_content:
            self._last_content = current_content
            return True
        return False
    
    def get_last_content(self) -> Optional[str]:
        return self._last_content
