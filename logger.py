import logging
import os
from datetime import datetime
from typing import Optional


class Logger:
    def __init__(self, log_dir: str = "./logs", log_level: int = logging.INFO):
        self.log_dir = log_dir
        self.log_level = log_level
        self.logger = self._setup_logger()
    
    def _setup_logger(self) -> logging.Logger:
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        
        logger = logging.getLogger("AIword")
        logger.setLevel(self.log_level)
        
        if not logger.handlers:
            log_file = os.path.join(self.log_dir, f"{datetime.now().strftime('%Y-%m-%d')}.log")
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setLevel(self.log_level)
            
            console_handler = logging.StreamHandler()
            console_handler.setLevel(self.log_level)
            
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            file_handler.setFormatter(formatter)
            console_handler.setFormatter(formatter)
            
            logger.addHandler(file_handler)
            logger.addHandler(console_handler)
        
        return logger
    
    def info(self, message: str) -> None:
        self.logger.info(message)
    
    def warning(self, message: str) -> None:
        self.logger.warning(message)
    
    def error(self, message: str) -> None:
        self.logger.error(message)
    
    def debug(self, message: str) -> None:
        self.logger.debug(message)
    
    def get_log_path(self) -> str:
        return self.log_dir
