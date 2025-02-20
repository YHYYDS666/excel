import tkinter as tk
from tkinter import ttk

class StatusManager:
    """状态栏管理器"""
    def __init__(self, root):
        self.root = root
        self.progress_var = tk.IntVar()
        self.create_statusbar()
        
    def create_statusbar(self):
        """创建状态栏"""
        self.statusbar = ttk.Frame(self.root)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = ttk.Label(self.statusbar, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        self.progress = ttk.Progressbar(
            self.statusbar,
            mode='determinate',
            variable=self.progress_var
        )
        self.progress.pack(side=tk.RIGHT, padx=5)
        
    def update_status(self, text: str, progress: int = None):
        """更新状态
        Args:
            text: 状态文本
            progress: 进度值(0-100)
        """
        self.status_label.config(text=text)
        if progress is not None:
            self.progress_var.set(progress) 