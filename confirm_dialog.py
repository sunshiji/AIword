import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any, List
from enum import Enum
from logger import Logger


class ConfirmAction(Enum):
    CONFIRM = "confirm"
    COPY_TO_CLIPBOARD = "copy_to_clipboard"
    SAVE_FILE = "save_file"
    CANCEL = "cancel"


class ContentConfirmDialog:
    def __init__(self, logger: Logger):
        self.logger = logger
        self.result = None
    
    def show(self, 
             content: str,
             content_type: str,
             html_content: str = None,
             csv_content: str = None,
             word_available: bool = False,
             excel_available: bool = False,
             no_app_action: str = "copy_to_clipboard") -> ConfirmAction:
        
        self.result = None
        self.word_available = word_available
        self.excel_available = excel_available
        self.no_app_action = no_app_action
        self.content = content
        self.content_type = content_type
        self.html_content = html_content
        self.csv_content = csv_content
        
        self.window = tk.Tk()
        self.window.title("AIword - 确认粘贴")
        self.window.geometry("600x500")
        self.window.resizable(True, True)
        
        self._create_widgets()
        
        self.window.eval('tk::PlaceWindow . center')
        self.window.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.window.grab_set()
        self.window.focus_force()
        self.window.mainloop()
        
        return self.result or ConfirmAction.CANCEL
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            title_frame, 
            text="AIword - 粘贴确认", 
            font=("TkDefaultFont", 12, "bold")
        ).pack(side=tk.LEFT)
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        type_label = self._get_content_type_label()
        ttk.Label(
            status_frame, 
            text=f"内容类型: {type_label}",
            font=("TkDefaultFont", 9)
        ).pack(anchor=tk.W)
        
        app_status = []
        if self.content_type == 'markdown_table':
            if self.excel_available:
                app_status.append("✓ Excel/WPS 表格 已运行")
            else:
                app_status.append("✗ Excel/WPS 表格 未运行")
        else:
            if self.word_available:
                app_status.append("✓ Word/WPS 已运行")
            else:
                app_status.append("✗ Word/WPS 未运行")
        
        ttk.Label(
            status_frame, 
            text=" | ".join(app_status),
            font=("TkDefaultFont", 9)
        ).pack(anchor=tk.W)
        
        preview_frame = ttk.LabelFrame(main_frame, text="内容预览", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        preview_text = tk.Text(
            preview_frame, 
            wrap=tk.WORD, 
            font=("Consolas", 9),
            padx=5,
            pady=5
        )
        preview_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=preview_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        preview_text.configure(yscrollcommand=scrollbar.set)
        
        preview_text.insert(tk.END, self._get_preview_text())
        preview_text.configure(state=tk.DISABLED)
        
        action_frame = ttk.LabelFrame(main_frame, text="选择操作", padding="10")
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        action_inner = ttk.Frame(action_frame)
        action_inner.pack(fill=tk.X)
        
        self.action_var = tk.StringVar(value=self._get_default_action())
        
        actions = self._get_available_actions()
        
        for i, (label, value) in enumerate(actions):
            ttk.Radiobutton(
                action_inner, 
                text=label, 
                value=value, 
                variable=self.action_var
            ).grid(row=i // 2, column=i % 2, sticky=tk.W, padx=5, pady=2)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(
            button_frame, 
            text="确定", 
            command=self._on_confirm,
            style="Accent.TButton"
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="取消", 
            command=self._on_cancel
        ).pack(side=tk.RIGHT, padx=5)
    
    def _get_content_type_label(self) -> str:
        type_map = {
            'markdown_table': 'Markdown 表格',
            'html': 'HTML 富文本',
            'plain_text': '纯文本'
        }
        return type_map.get(self.content_type, self.content_type)
    
    def _get_preview_text(self) -> str:
        if self.content_type == 'markdown_table' and self.csv_content:
            return self.csv_content.replace(',', '\t')
        else:
            preview = self.content
            if len(preview) > 5000:
                preview = preview[:5000] + "\n...(内容已截断)"
            return preview
    
    def _get_default_action(self) -> str:
        if self.content_type == 'markdown_table':
            if self.excel_available:
                return 'confirm'
        else:
            if self.word_available:
                return 'confirm'
        
        action_map = {
            'auto_open': 'confirm',
            'save_only': 'save_file',
            'copy_to_clipboard': 'copy_to_clipboard',
            'no_action': 'cancel'
        }
        return action_map.get(self.no_app_action, 'copy_to_clipboard')
    
    def _get_available_actions(self) -> List[tuple]:
        actions = []
        
        if self.content_type == 'markdown_table':
            if self.excel_available:
                actions.append(("粘贴到 Excel/WPS 表格", "confirm"))
            actions.append(("打开 Excel/WPS 表格并粘贴", "confirm"))
        else:
            if self.word_available:
                actions.append(("粘贴到 Word/WPS", "confirm"))
            actions.append(("打开 Word/WPS 并粘贴", "confirm"))
        
        actions.extend([
            ("复制到剪贴板", "copy_to_clipboard"),
            ("保存为文件", "save_file"),
        ])
        
        return actions
    
    def _on_confirm(self):
        action = self.action_var.get()
        
        action_map = {
            'confirm': ConfirmAction.CONFIRM,
            'copy_to_clipboard': ConfirmAction.COPY_TO_CLIPBOARD,
            'save_file': ConfirmAction.SAVE_FILE,
            'cancel': ConfirmAction.CANCEL
        }
        
        self.result = action_map.get(action, ConfirmAction.CANCEL)
        self.window.destroy()
    
    def _on_cancel(self):
        self.result = ConfirmAction.CANCEL
        self.window.destroy()


def show_confirm_dialog(logger: Logger,
                        content: str,
                        content_type: str,
                        html_content: str = None,
                        csv_content: str = None,
                        word_available: bool = False,
                        excel_available: bool = False,
                        no_app_action: str = "copy_to_clipboard") -> ConfirmAction:
    dialog = ContentConfirmDialog(logger)
    return dialog.show(
        content=content,
        content_type=content_type,
        html_content=html_content,
        csv_content=csv_content,
        word_available=word_available,
        excel_available=excel_available,
        no_app_action=no_app_action
    )
