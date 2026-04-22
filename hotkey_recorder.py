import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable
import keyboard
from logger import Logger


class HotkeyRecorder:
    MODIFIER_KEYS = {
        'ctrl': 'ctrl',
        'control': 'ctrl',
        'alt': 'alt',
        'shift': 'shift',
        'win': 'win',
        'windows': 'win'
    }
    
    def __init__(self, logger: Logger, current_hotkey: str = "ctrl+alt+v"):
        self.logger = logger
        self.current_hotkey = current_hotkey
        self.recording = False
        self.recorded_keys = set()
        self.result = None
        self.window = None
        self.lbl_hotkey = None
        self.btn_record = None
    
    def normalize_key(self, key: str) -> str:
        key_lower = key.lower()
        if key_lower in self.MODIFIER_KEYS:
            return self.MODIFIER_KEYS[key_lower]
        return key
    
    def start_recording(self):
        self.recording = True
        self.recorded_keys.clear()
        self.btn_record.config(text="停止录制", style="Accent.TButton")
        self.lbl_hotkey.config(text="请按下组合键...")
        self.logger.info("开始录制热键")
    
    def stop_recording(self):
        self.recording = False
        self.btn_record.config(text="录制热键", style="TButton")
        
        if self.recorded_keys:
            hotkey_parts = sorted(self.recorded_keys, key=lambda k: (k not in ['ctrl', 'alt', 'shift', 'win'], k))
            hotkey_str = '+'.join(hotkey_parts)
            self.lbl_hotkey.config(text=f"当前热键: {hotkey_str}")
            self.result = hotkey_str
            self.logger.info(f"录制热键: {hotkey_str}")
        else:
            self.lbl_hotkey.config(text=f"当前热键: {self.current_hotkey}")
            self.result = None
    
    def on_key_event(self, event):
        if not self.recording:
            return
        
        if event.event_type == 'down':
            key = self.normalize_key(event.name)
            self.recorded_keys.add(key)
            
            hotkey_parts = sorted(self.recorded_keys, key=lambda k: (k not in ['ctrl', 'alt', 'shift', 'win'], k))
            hotkey_str = '+'.join(hotkey_parts)
            self.lbl_hotkey.config(text=f"正在录制: {hotkey_str}")
    
    def show_dialog(self, parent: Optional[tk.Tk] = None) -> Optional[str]:
        self.result = None
        
        if parent:
            self.window = tk.Toplevel(parent)
        else:
            self.window = tk.Tk()
        
        self.window.title("设置热键")
        self.window.geometry("350x200")
        self.window.resizable(False, False)
        
        if not parent:
            self.window.eval('tk::PlaceWindow . center')
        else:
            self.window.transient(parent)
            self.window.grab_set()
        
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="录制新的全局热键", font=("TkDefaultFont", 10, "bold")).pack(pady=(0, 10))
        
        self.lbl_hotkey = ttk.Label(
            main_frame, 
            text=f"当前热键: {self.current_hotkey}",
            font=("TkDefaultFont", 12)
        )
        self.lbl_hotkey.pack(pady=10)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        self.btn_record = ttk.Button(btn_frame, text="录制热键", command=self.toggle_recording)
        self.btn_record.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="确定", command=self.confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        keyboard.hook(self.on_key_event)
        
        self.window.protocol("WM_DELETE_WINDOW", self.cancel)
        
        if not parent:
            self.window.mainloop()
        
        keyboard.unhook(self.on_key_event)
        
        return self.result
    
    def toggle_recording(self):
        if self.recording:
            self.stop_recording()
        else:
            self.start_recording()
    
    def confirm(self):
        if self.result:
            self.window.destroy()
        else:
            messagebox.showinfo("提示", "请先录制一个热键")
    
    def cancel(self):
        self.result = None
        if self.recording:
            self.stop_recording()
        self.window.destroy()


def record_hotkey(logger: Logger, current_hotkey: str = "ctrl+alt+v") -> Optional[str]:
    recorder = HotkeyRecorder(logger, current_hotkey)
    return recorder.show_dialog()
