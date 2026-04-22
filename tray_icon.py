import os
import threading
from typing import Optional, Callable
from PIL import Image
import pystray
from pystray import MenuItem as item

from config_manager import ConfigManager
from logger import Logger
from hotkey_manager import HotkeyManager
from update_checker import UpdateChecker
from settings_window import show_settings


class TrayIcon:
    def __init__(self, 
                 config_manager: ConfigManager, 
                 logger: Logger,
                 hotkey_manager: HotkeyManager,
                 on_exit: Optional[Callable] = None,
                 on_process: Optional[Callable] = None):
        self.config_manager = config_manager
        self.logger = logger
        self.hotkey_manager = hotkey_manager
        self.on_exit = on_exit
        self.on_process = on_process
        self.icon = None
        self.update_checker = UpdateChecker(
            logger, 
            config_manager.get('version', '1.0.0')
        )
        self.has_update = False
    
    def _create_icon_image(self):
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'icon.png')
            if os.path.exists(icon_path):
                return Image.open(icon_path)
        except Exception:
            pass
        
        from io import BytesIO
        img = Image.new('RGB', (64, 64), color='#3498db')
        return img
    
    def _build_menu(self):
        menu_items = []
        
        if self.on_process:
            menu_items.append(item('粘贴内容', self._on_paste, default=True))
            menu_items.append(item('', None))
        
        version = self.config_manager.get('version', '1.0.0')
        if self.has_update:
            latest_ver = self.update_checker.get_latest_version() or '新版'
            menu_items.append(item(f'有新版本: v{latest_ver}', self._open_download))
        else:
            menu_items.append(item(f'版本: v{version}', None, enabled=False))
        
        menu_items.append(item('检查更新', self._check_update))
        menu_items.append(item('', None))
        
        hotkey_status = "启用" if self.hotkey_manager.is_enabled() else "禁用"
        current_hotkey = self.hotkey_manager.get_current_hotkey()
        menu_items.append(item(f'热键: {current_hotkey} ({hotkey_status})', None, enabled=False))
        menu_items.append(item('启用热键', self._toggle_hotkey, checked=lambda item: self.hotkey_manager.is_enabled()))
        menu_items.append(item('设置热键', self._set_hotkey))
        menu_items.append(item('', None))
        
        menu_items.append(item('设置', self._show_settings))
        menu_items.append(item('重载配置', self._reload_config))
        menu_items.append(item('', None))
        menu_items.append(item('退出', self._on_exit))
        
        return tuple(menu_items)
    
    def _update_menu(self):
        if self.icon:
            self.icon.menu = self._build_menu()
    
    def _on_paste(self, icon, item):
        if self.on_process:
            threading.Thread(target=self.on_process, daemon=True).start()
    
    def _on_exit(self, icon, item):
        self.logger.info("正在退出程序...")
        if self.icon:
            self.icon.stop()
        if self.on_exit:
            self.on_exit()
    
    def _show_settings(self, icon, item):
        threading.Thread(
            target=lambda: show_settings(self.config_manager, self.logger, self._on_config_changed),
            daemon=True
        ).start()
    
    def _toggle_hotkey(self, icon, item):
        if self.hotkey_manager.is_enabled():
            self.hotkey_manager.disable()
            self.config_manager.set('hotkey_enabled', False)
        else:
            self.hotkey_manager.enable()
            self.config_manager.set('hotkey_enabled', True)
        self._update_menu()
        self.logger.info(f"热键已{'启用' if self.hotkey_manager.is_enabled() else '禁用'}")
    
    def _set_hotkey(self, icon, item):
        from hotkey_recorder import HotkeyRecorder
        
        def record():
            current_hotkey = self.hotkey_manager.get_current_hotkey()
            recorder = HotkeyRecorder(self.logger, current_hotkey)
            new_hotkey = recorder.show_dialog()
            
            if new_hotkey:
                self.hotkey_manager.set_hotkey(new_hotkey)
                self.config_manager.set('hotkey', new_hotkey)
                self._update_menu()
        
        threading.Thread(target=record, daemon=True).start()
    
    def _reload_config(self, icon, item):
        self.config_manager.reload()
        self.hotkey_manager.set_hotkey(self.config_manager.get('hotkey', 'ctrl+alt+v'))
        
        if self.config_manager.get('hotkey_enabled', True):
            self.hotkey_manager.enable()
        else:
            self.hotkey_manager.disable()
        
        self._update_menu()
        self.logger.info("配置已重新加载")
    
    def _check_update(self, icon, item):
        def check():
            self.has_update = self.update_checker.check_for_updates()
            self._update_menu()
            
            if self.has_update:
                latest_ver = self.update_checker.get_latest_version()
                self.logger.info(f"发现新版本: v{latest_ver}")
            else:
                self.logger.info("当前已是最新版本")
        
        threading.Thread(target=check, daemon=True).start()
    
    def _open_download(self, icon, item):
        import webbrowser
        url = self.update_checker.get_download_url()
        if url:
            webbrowser.open(url)
            self.logger.info(f"正在打开下载页面: {url}")
    
    def _on_config_changed(self):
        self.hotkey_manager.set_hotkey(self.config_manager.get('hotkey', 'ctrl+alt+v'))
        
        if self.config_manager.get('hotkey_enabled', True):
            self.hotkey_manager.enable()
        else:
            self.hotkey_manager.disable()
        
        self._update_menu()
    
    def run(self):
        self.logger.info("启动系统托盘图标...")
        
        image = self._create_icon_image()
        menu = self._build_menu()
        
        self.icon = pystray.Icon(
            name='AIword',
            icon=image,
            title='AIword - 智能粘贴工具',
            menu=menu
        )
        
        self.icon.run()
    
    def stop(self):
        if self.icon:
            self.icon.stop()
