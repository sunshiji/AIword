import requests
import re
from typing import Optional, Tuple
from logger import Logger


class UpdateChecker:
    GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/releases/latest"
    DEFAULT_REPO = "user/AIword"
    
    def __init__(self, logger: Logger, current_version: str, repo: str = None):
        self.logger = logger
        self.current_version = current_version
        self.repo = repo or self.DEFAULT_REPO
        self.latest_version = None
        self.download_url = None
        self.release_notes = None
    
    def parse_version(self, version_str: str) -> Tuple[int, ...]:
        match = re.search(r'(\d+)\.(\d+)\.(\d+)', version_str)
        if match:
            return tuple(map(int, match.groups()))
        return (0, 0, 0)
    
    def compare_versions(self, v1: str, v2: str) -> int:
        ver1 = self.parse_version(v1)
        ver2 = self.parse_version(v2)
        
        if ver1 > ver2:
            return 1
        elif ver1 < ver2:
            return -1
        else:
            return 0
    
    def check_for_updates(self) -> bool:
        try:
            owner, repo = self.repo.split('/')
            url = self.GITHUB_API_URL.format(owner=owner, repo=repo)
            
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            
            self.latest_version = data.get('tag_name', 'v0.0.0').lstrip('v')
            self.release_notes = data.get('body', '')
            
            assets = data.get('assets', [])
            if assets:
                self.download_url = assets[0].get('browser_download_url')
            else:
                self.download_url = data.get('html_url')
            
            comparison = self.compare_versions(self.latest_version, self.current_version)
            
            if comparison > 0:
                self.logger.info(f"发现新版本: {self.latest_version} (当前版本: {self.current_version})")
                return True
            else:
                self.logger.info(f"当前已是最新版本: {self.current_version}")
                return False
                
        except Exception as e:
            self.logger.error(f"检查更新失败: {str(e)}")
            return False
    
    def get_latest_version(self) -> Optional[str]:
        return self.latest_version
    
    def get_download_url(self) -> Optional[str]:
        return self.download_url
    
    def get_release_notes(self) -> Optional[str]:
        return self.release_notes
    
    def has_update(self) -> bool:
        return self.latest_version is not None and self.compare_versions(self.latest_version, self.current_version) > 0
