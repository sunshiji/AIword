import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import subprocess
from typing import Optional, Callable, Dict, Any

from config_manager import ConfigManager
from logger import Logger
from hotkey_recorder import HotkeyRecorder


class SettingsWindow:
    NO_APP_ACTIONS = [
        ("自动打开", "auto_open"),
        ("仅保存", "save_only"),
        ("复制到剪贴板", "copy_to_clipboard"),
        ("无操作", "no_action")
    ]
    
    def __init__(self, config_manager: ConfigManager, logger: Logger, 
                 on_config_changed: Optional[Callable] = None):
        self.config_manager = config_manager
        self.logger = logger
        self.on_config_changed = on_config_changed
        self.window = None
        self.vars: Dict[str, Any] = {}
    
    def show(self, parent: Optional[tk.Tk] = None):
        if parent:
            self.window = tk.Toplevel(parent)
        else:
            self.window = tk.Tk()
        
        self.window.title("AIword 设置")
        self.window.geometry("500x550")
        self.window.resizable(False, False)
        
        if not parent:
            self.window.eval('tk::PlaceWindow . center')
        else:
            self.window.transient(parent)
            self.window.grab_set()
        
        self._create_widgets()
        self._load_settings()
        
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)
        
        if not parent:
            self.window.mainloop()
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="AIword 设置", font=("TkDefaultFont", 14, "bold")).pack(pady=(0, 15))
        
        hotkey_frame = ttk.LabelFrame(main_frame, text="热键设置", padding="10")
        hotkey_frame.pack(fill=tk.X, pady=5)
        
        hotkey_row = ttk.Frame(hotkey_frame)
        hotkey_row.pack(fill=tk.X)
        
        self.vars['current_hotkey'] = tk.StringVar()
        ttk.Label(hotkey_row, textvariable=self.vars['current_hotkey']).pack(side=tk.LEFT, padx=5)
        ttk.Button(hotkey_row, text="设置热键", command=self._set_hotkey).pack(side=tk.LEFT, padx=5)
        
        self.vars['hotkey_enabled'] = tk.BooleanVar()
        ttk.Checkbutton(hotkey_frame, text="启用全局热键", variable=self.vars['hotkey_enabled']).pack(anchor=tk.W, pady=5)
        
        options_frame = ttk.LabelFrame(main_frame, text="选项设置", padding="10")
        options_frame.pack(fill=tk.X, pady=5)
        
        self.vars['notification_enabled'] = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="启用弹窗通知", variable=self.vars['notification_enabled']).pack(anchor=tk.W)
        
        self.vars['move_cursor_to_end'] = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="插入后移动光标到末尾", variable=self.vars['move_cursor_to_end']).pack(anchor=tk.W)
        
        self.vars['html_formatting'] = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="HTML 格式化（Markdown 样式转换）", variable=self.vars['html_formatting']).pack(anchor=tk.W)
        
        self.vars['keep_generated_files'] = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="保留生成文件", variable=self.vars['keep_generated_files']).pack(anchor=tk.W)
        
        self.vars['confirm_before_paste'] = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="粘贴前确认（显示预览对话框）", variable=self.vars['confirm_before_paste']).pack(anchor=tk.W)
        
        action_frame = ttk.LabelFrame(main_frame, text="无应用时动作", padding="10")
        action_frame.pack(fill=tk.X, pady=5)
        
        self.vars['no_app_action'] = tk.StringVar()
        action_combo = ttk.Combobox(action_frame, textvariable=self.vars['no_app_action'], 
                                     values=[text for text, value in self.NO_APP_ACTIONS],
                                     state="readonly", width=25)
        action_combo.pack(anchor=tk.W)
        action_combo.bind('<<ComboboxSelected>>', self._on_action_changed)
        
        dir_frame = ttk.LabelFrame(main_frame, text="保存目录", padding="10")
        dir_frame.pack(fill=tk.X, pady=5)
        
        dir_row = ttk.Frame(dir_frame)
        dir_row.pack(fill=tk.X)
        
        self.vars['save_dir'] = tk.StringVar()
        ttk.Entry(dir_row, textvariable=self.vars['save_dir'], width=40).pack(side=tk.LEFT, padx=5)
        ttk.Button(dir_row, text="浏览...", command=self._browse_save_dir).pack(side=tk.LEFT)
        
        tools_frame = ttk.LabelFrame(main_frame, text="工具", padding="10")
        tools_frame.pack(fill=tk.X, pady=5)
        
        tools_row = ttk.Frame(tools_frame)
        tools_row.pack(fill=tk.X)
        
        ttk.Button(tools_row, text="打开保存目录", command=self._open_save_dir).pack(side=tk.LEFT, padx=5)
        ttk.Button(tools_row, text="查看日志", command=self._view_logs).pack(side=tk.LEFT, padx=5)
        ttk.Button(tools_row, text="编辑配置", command=self._edit_config).pack(side=tk.LEFT, padx=5)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(button_frame, text="保存", command=self._save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="重置", command=self._reset_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关闭", command=self._on_close).pack(side=tk.RIGHT, padx=5)
    
    def _load_settings(self):
        config = self.config_manager.get_all()
        
        self.vars['current_hotkey'].set(f"当前全局热键: {config.get('hotkey', 'ctrl+alt+v')}")
        self.vars['hotkey_enabled'].set(config.get('hotkey_enabled', True))
        self.vars['notification_enabled'].set(config.get('notification_enabled', True))
        self.vars['move_cursor_to_end'].set(config.get('move_cursor_to_end', True))
        self.vars['html_formatting'].set(config.get('html_formatting', True))
        self.vars['keep_generated_files'].set(config.get('keep_generated_files', False))
        self.vars['confirm_before_paste'].set(config.get('confirm_before_paste', True))
        self.vars['save_dir'].set(config.get('save_dir', './output'))
        
        action_value = config.get('no_app_action', 'copy_to_clipboard')
        for text, value in self.NO_APP_ACTIONS:
            if value == action_value:
                self.vars['no_app_action'].set(text)
                break
    
    def _set_hotkey(self):
        current_hotkey = self.config_manager.get('hotkey', 'ctrl+alt+v')
        recorder = HotkeyRecorder(self.logger, current_hotkey)
        new_hotkey = recorder.show_dialog(self.window)
        
        if new_hotkey:
            self.vars['current_hotkey'].set(f"当前全局热键: {new_hotkey}")
            self.config_manager.set('hotkey', new_hotkey)
    
    def _on_action_changed(self, event=None):
        pass
    
    def _browse_save_dir(self):
        current_dir = self.vars['save_dir'].get()
        if not os.path.isabs(current_dir):
            current_dir = os.path.abspath(current_dir)
        
        directory = filedialog.askdirectory(
            title="选择保存目录",
            initialdir=current_dir if os.path.exists(current_dir) else None
        )
        
        if directory:
            self.vars['save_dir'].set(directory)
    
    def _open_save_dir(self):
        save_dir = self.vars['save_dir'].get()
        if not os.path.isabs(save_dir):
            save_dir = os.path.abspath(save_dir)
        
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        try:
            os.startfile(save_dir)
        except Exception as e:
            self.logger.error(f"打开目录失败: {str(e)}")
            messagebox.showerror("错误", f"打开目录失败: {str(e)}")
    
    def _view_logs(self):
        log_dir = self.config_manager.get('log_dir', './logs')
        if not os.path.isabs(log_dir):
            log_dir = os.path.abspath(log_dir)
        
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        try:
            os.startfile(log_dir)
        except Exception as e:
            self.logger.error(f"打开日志目录失败: {str(e)}")
            messagebox.showerror("错误", f"打开日志目录失败: {str(e)}")
    
    def _edit_config(self):
        config_path = self.config_manager.config_path
        if not os.path.isabs(config_path):
            config_path = os.path.abspath(config_path)
        
        try:
            os.startfile(config_path)
        except Exception as e:
            self.logger.error(f"打开配置文件失败: {str(e)}")
            messagebox.showerror("错误", f"打开配置文件失败: {str(e)}")
    
    def _save_settings(self):
        action_text = self.vars['no_app_action'].get()
        action_value = 'copy_to_clipboard'
        for text, value in self.NO_APP_ACTIONS:
            if text == action_text:
                action_value = value
                break
        
        self.config_manager.set('hotkey_enabled', self.vars['hotkey_enabled'].get())
        self.config_manager.set('notification_enabled', self.vars['notification_enabled'].get())
        self.config_manager.set('move_cursor_to_end', self.vars['move_cursor_to_end'].get())
        self.config_manager.set('html_formatting', self.vars['html_formatting'].get())
        self.config_manager.set('keep_generated_files', self.vars['keep_generated_files'].get())
        self.config_manager.set('confirm_before_paste', self.vars['confirm_before_paste'].get())
        self.config_manager.set('no_app_action', action_value)
        self.config_manager.set('save_dir', self.vars['save_dir'].get())
        
        messagebox.showinfo("成功", "设置已保存")
        
        if self.on_config_changed:
            self.on_config_changed()
    
    def _reset_settings(self):
        self._load_settings()
    
    def _on_close(self):
        self.window.destroy()


def show_settings(config_manager: ConfigManager, logger: Logger, 
                  on_config_changed: Optional[Callable] = None):
    settings = SettingsWindow(config_manager, logger, on_config_changed)
    settings.show()
