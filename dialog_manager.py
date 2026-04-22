import tkinter as tk
from tkinter import ttk
from typing import Optional, Dict, Any, List, Callable
from queue import Queue, Empty
from enum import Enum
from dataclasses import dataclass
from logger import Logger


class DialogType(Enum):
    CONFIRM_PASTE = "confirm_paste"
    MESSAGE = "message"
    ERROR = "error"


class ConfirmAction(Enum):
    CONFIRM = "confirm"
    COPY_TO_CLIPBOARD = "copy_to_clipboard"
    SAVE_FILE = "save_file"
    CANCEL = "cancel"


@dataclass
class DialogRequest:
    dialog_type: DialogType
    data: Dict[str, Any]
    result_queue: Optional[Queue] = None


class DialogManager:
    _instance = None
    
    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, logger: Optional[Logger] = None):
        if not hasattr(self, '_initialized'):
            self._initialized = True
            self.logger = logger
            self.root = None
            self.request_queue = Queue()
            self.dialog_active = False
            self._after_id = None
    
    def initialize(self, root: tk.Tk):
        self.root = root
        self._start_polling()
        if self.logger:
            self.logger.info("对话框管理器已初始化")
    
    def _start_polling(self):
        try:
            while True:
                try:
                    request = self.request_queue.get_nowait()
                    self._handle_request(request)
                except Empty:
                    break
        except Exception as e:
            if self.logger:
                self.logger.error(f"处理对话框请求时出错: {str(e)}")
        
        if self.root:
            self._after_id = self.root.after(100, self._start_polling)
    
    def _handle_request(self, request: DialogRequest):
        if request.dialog_type == DialogType.CONFIRM_PASTE:
            self._show_confirm_dialog(request)
        elif request.dialog_type == DialogType.MESSAGE:
            self._show_message(request)
        elif request.dialog_type == DialogType.ERROR:
            self._show_error(request)
    
    def request_confirm_paste(self,
                               content: str,
                               content_type: str,
                               html_content: str = None,
                               csv_content: str = None,
                               word_available: bool = False,
                               excel_available: bool = False,
                               no_app_action: str = "copy_to_clipboard") -> ConfirmAction:
        result_queue = Queue()
        
        request = DialogRequest(
            dialog_type=DialogType.CONFIRM_PASTE,
            data={
                'content': content,
                'content_type': content_type,
                'html_content': html_content,
                'csv_content': csv_content,
                'word_available': word_available,
                'excel_available': excel_available,
                'no_app_action': no_app_action
            },
            result_queue=result_queue
        )
        
        self.request_queue.put(request)
        
        try:
            result = result_queue.get(timeout=300)
            return result
        except:
            return ConfirmAction.CANCEL
    
    def _show_confirm_dialog(self, request: DialogRequest):
        data = request.data
        
        dialog = tk.Toplevel(self.root)
        dialog.title("AIword - 确认粘贴")
        dialog.geometry("600x500")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        result = [ConfirmAction.CANCEL]
        action_var = tk.StringVar()
        
        self._setup_confirm_dialog_widgets(
            dialog,
            data,
            action_var,
            result
        )
        
        def on_close():
            dialog.destroy()
        
        dialog.protocol("WM_DELETE_WINDOW", on_close)
        dialog.eval('tk::PlaceWindow . center')
        
        dialog.wait_window()
        
        if request.result_queue:
            request.result_queue.put(result[0])
    
    def _setup_confirm_dialog_widgets(self, dialog, data, action_var, result):
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            main_frame, 
            text="AIword - 粘贴确认", 
            font=("TkDefaultFont", 12, "bold")
        ).pack(pady=(0, 10))
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        type_map = {
            'markdown_table': 'Markdown 表格',
            'html': 'HTML 富文本',
            'plain_text': '纯文本'
        }
        type_label = type_map.get(data['content_type'], data['content_type'])
        
        ttk.Label(
            status_frame, 
            text=f"内容类型: {type_label}",
            font=("TkDefaultFont", 9)
        ).pack(anchor=tk.W)
        
        app_status = []
        if data['content_type'] == 'markdown_table':
            if data['excel_available']:
                app_status.append("✓ Excel/WPS 表格 已运行")
            else:
                app_status.append("✗ Excel/WPS 表格 未运行")
        else:
            if data['word_available']:
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
        preview_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=preview_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        preview_text.configure(yscrollcommand=scrollbar.set)
        
        preview_content = data['content']
        if data['content_type'] == 'markdown_table' and data.get('csv_content'):
            preview_content = data['csv_content'].replace(',', '\t')
        
        if len(preview_content) > 5000:
            preview_content = preview_content[:5000] + "\n...(内容已截断)"
        
        preview_text.insert(tk.END, preview_content)
        preview_text.configure(state=tk.DISABLED)
        
        action_frame = ttk.LabelFrame(main_frame, text="选择操作", padding="10")
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        action_inner = ttk.Frame(action_frame)
        action_inner.pack(fill=tk.X)
        
        default_action = self._get_default_action(data)
        action_var.set(default_action)
        
        actions = self._get_available_actions(data)
        
        for i, (label, value) in enumerate(actions):
            ttk.Radiobutton(
                action_inner, 
                text=label, 
                value=value, 
                variable=action_var
            ).grid(row=i // 2, column=i % 2, sticky=tk.W, padx=5, pady=2)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def on_confirm():
            action = action_var.get()
            action_map = {
                'confirm': ConfirmAction.CONFIRM,
                'copy_to_clipboard': ConfirmAction.COPY_TO_CLIPBOARD,
                'save_file': ConfirmAction.SAVE_FILE,
                'cancel': ConfirmAction.CANCEL
            }
            result[0] = action_map.get(action, ConfirmAction.CANCEL)
            dialog.destroy()
        
        def on_cancel():
            result[0] = ConfirmAction.CANCEL
            dialog.destroy()
        
        ttk.Button(
            button_frame, 
            text="确定", 
            command=on_confirm,
            style="Accent.TButton"
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="取消", 
            command=on_cancel
        ).pack(side=tk.RIGHT, padx=5)
    
    def _get_default_action(self, data) -> str:
        if data['content_type'] == 'markdown_table':
            if data['excel_available']:
                return 'confirm'
        else:
            if data['word_available']:
                return 'confirm'
        
        action_map = {
            'auto_open': 'confirm',
            'save_only': 'save_file',
            'copy_to_clipboard': 'copy_to_clipboard',
            'no_action': 'cancel'
        }
        return action_map.get(data.get('no_app_action', 'copy_to_clipboard'), 'copy_to_clipboard')
    
    def _get_available_actions(self, data) -> List[tuple]:
        actions = []
        
        if data['content_type'] == 'markdown_table':
            if data['excel_available']:
                actions.append(("粘贴到 Excel/WPS 表格", "confirm"))
            actions.append(("打开 Excel/WPS 表格并粘贴", "confirm"))
        else:
            if data['word_available']:
                actions.append(("粘贴到 Word/WPS", "confirm"))
            actions.append(("打开 Word/WPS 并粘贴", "confirm"))
        
        actions.extend([
            ("复制到剪贴板", "copy_to_clipboard"),
            ("保存为文件", "save_file"),
        ])
        
        return actions
    
    def request_message(self, title: str, message: str):
        request = DialogRequest(
            dialog_type=DialogType.MESSAGE,
            data={'title': title, 'message': message}
        )
        self.request_queue.put(request)
    
    def _show_message(self, request: DialogRequest):
        from tkinter import messagebox
        if self.root:
            self.root.after(0, lambda: messagebox.showinfo(
                request.data.get('title', '提示'),
                request.data.get('message', '')
            ))
    
    def request_error(self, title: str, message: str):
        request = DialogRequest(
            dialog_type=DialogType.ERROR,
            data={'title': title, 'message': message}
        )
        self.request_queue.put(request)
    
    def _show_error(self, request: DialogRequest):
        from tkinter import messagebox
        if self.root:
            self.root.after(0, lambda: messagebox.showerror(
                request.data.get('title', '错误'),
                request.data.get('message', '')
            ))
    
    def stop(self):
        if self._after_id and self.root:
            try:
                self.root.after_cancel(self._after_id)
            except:
                pass
            self._after_id = None


dialog_manager: Optional[DialogManager] = None


def get_dialog_manager(logger: Optional[Logger] = None) -> DialogManager:
    global dialog_manager
    if dialog_manager is None:
        dialog_manager = DialogManager(logger)
    return dialog_manager


def initialize_dialog_manager(root: tk.Tk, logger: Optional[Logger] = None):
    dm = get_dialog_manager(logger)
    dm.initialize(root)
