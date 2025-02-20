import tkinter as tk
from tkinter import ttk

class ToolbarManager:
    """工具栏管理器"""
    def __init__(self, root):
        self.root = root
        self.toolbar = None
        self.create_toolbar()
        
    def create_toolbar(self):
        """创建工具栏"""
        # 如果已存在工具栏，先销毁
        if self.toolbar:
            self.toolbar.destroy()
            
        self.toolbar = ttk.Frame(self.root)
        self.toolbar.pack(fill=tk.X, padx=10, pady=5)
        
        # 创建自定义按钮样式
        style = ttk.Style()
        style.configure('Accent.TButton', 
            background='#0078D7',  # Windows蓝色
            foreground='white'
        )
        
        # 文件操作按钮组
        file_frame = ttk.Frame(self.toolbar)
        file_frame.pack(side=tk.LEFT)
        
        self.open_btn = ttk.Button(
            file_frame,
            text="打开Excel",
            width=10
        )
        self.open_btn.pack(side=tk.LEFT, padx=2)
        
        self.refresh_btn = ttk.Button(
            file_frame,
            text="刷新",
            width=8,
            state='disabled'  # 初始状态禁用
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=2)
        
        # 历史操作按钮组
        history_frame = ttk.Frame(self.toolbar)
        history_frame.pack(side=tk.LEFT, padx=20)
        
        self.undo_btn = ttk.Button(
            history_frame,
            text="撤销",
            width=8,
            state='disabled'
        )
        self.undo_btn.pack(side=tk.LEFT, padx=2)
        
        self.redo_btn = ttk.Button(
            history_frame,
            text="重做",
            width=8,
            state='disabled'
        )
        self.redo_btn.pack(side=tk.LEFT, padx=2)
        
        self.reset_btn = ttk.Button(
            history_frame,
            text="重置",
            width=8,
            state='disabled'
        )
        self.reset_btn.pack(side=tk.LEFT, padx=2)
        
        # 工作簿选择区域
        workbook_frame = ttk.Frame(self.toolbar)
        workbook_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=20)
        
        workbook_label = ttk.Label(workbook_frame, text="工作簿:")
        workbook_label.pack(side=tk.LEFT, padx=(0,5))
        
        self.workbook_var = tk.StringVar()
        self.workbook_combo = ttk.Combobox(
            workbook_frame,
            textvariable=self.workbook_var,
            state="disabled"  # 初始状态禁用
        )
        self.workbook_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
    def bind_commands(self, commands):
        """绑定按钮命令"""
        self.open_btn.config(command=commands.get('open'))
        self.refresh_btn.config(command=commands.get('refresh'))
        self.undo_btn.config(command=commands.get('undo'))
        self.redo_btn.config(command=commands.get('redo'))
        self.reset_btn.config(command=commands.get('reset'))
        
        # 绑定工作簿选择事件
        if 'workbook_selected' in commands:
            self.workbook_combo.bind('<<ComboboxSelected>>', 
                commands['workbook_selected'])

    def update_button_states(self, states):
        """更新按钮状态
        Args:
            states: 包含按钮状态的字典
        """
        if 'undo' in states:
            self.undo_btn['state'] = 'normal' if states['undo'] else 'disabled'
        if 'redo' in states:
            self.redo_btn['state'] = 'normal' if states['redo'] else 'disabled'
        if 'reset' in states:
            self.reset_btn['state'] = 'normal' if states['reset'] else 'disabled'

    def update_workbook_list(self, workbooks):
        """更新工作簿列表"""
        try:
            # 更新下拉框值
            self.workbook_combo['values'] = list(workbooks.keys())
            
            # 如果只有一个工作簿，自动选中
            if len(workbooks) == 1:
                self.workbook_var.set(list(workbooks.keys())[0])
            elif len(workbooks) == 0:
                self.workbook_var.set('')
            
            # 更新状态
            if workbooks:
                self.workbook_combo['state'] = 'readonly'
                self.refresh_btn['state'] = 'normal'
            else:
                self.workbook_combo['state'] = 'disabled'
                self.refresh_btn['state'] = 'disabled'
            
        except Exception as e:
            print(f"更新工作簿列表失败: {str(e)}")

    def update_excel_status(self, is_connected: bool):
        """更新Excel连接状态"""
        try:
            if is_connected:
                self.open_btn['state'] = 'normal'  # 改为始终可用
                self.refresh_btn['state'] = 'normal'
                self.workbook_combo['state'] = 'readonly'
            else:
                self.open_btn['state'] = 'normal'
                self.refresh_btn['state'] = 'disabled'
                self.workbook_combo['state'] = 'disabled'
                self.workbook_var.set('')
            
        except Exception as e:
            print(f"更新Excel状态失败: {str(e)}") 