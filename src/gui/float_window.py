import tkinter as tk
from tkinter import ttk

class FloatWindow(tk.Toplevel):
    """浮动窗口基类"""
    def __init__(self, parent, title="", width=300, height=200):
        super().__init__(parent)
        
        # 设置窗口属性
        self.title(title)
        self.geometry(f"{width}x{height}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # 居中显示
        self.center_window()
        
        # 创建界面
        self.create_widgets()
    
    def center_window(self):
        """窗口居中显示"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.geometry(f"+{x}+{y}")
    
    def create_widgets(self):
        """创建界面组件"""
        pass

    def create_sumif_widgets(self, parent):
        """创建条件求和设置界面"""
        # 条件类型
        type_frame = ttk.LabelFrame(parent, text="条件类型", padding=5)
        type_frame.pack(fill=tk.X, pady=(0,5))
        
        self.condition_type = tk.StringVar(value=">")
        ttk.Radiobutton(type_frame, text="大于", value=">", 
                       variable=self.condition_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="小于", value="<", 
                       variable=self.condition_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="等于", value="=", 
                       variable=self.condition_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="不等于", value="<>", 
                       variable=self.condition_type).pack(side=tk.LEFT, padx=5)
        
        # 条件值
        value_frame = ttk.LabelFrame(parent, text="条件值", padding=5)
        value_frame.pack(fill=tk.X)
        
        self.condition_value = ttk.Entry(value_frame)
        self.condition_value.pack(fill=tk.X, padx=5)
        
    def create_countif_widgets(self, parent):
        """创建条件计数设置界面"""
        # 计数类型
        type_frame = ttk.LabelFrame(parent, text="计数类型", padding=5)
        type_frame.pack(fill=tk.X, pady=(0,5))
        
        self.count_type = tk.StringVar(value="all")
        ttk.Radiobutton(type_frame, text="全部", value="all", 
                       variable=self.count_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="非空", value="nonblank", 
                       variable=self.count_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="数字", value="number", 
                       variable=self.count_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="文本", value="text", 
                       variable=self.count_type).pack(side=tk.LEFT, padx=5)
        
        # 条件值
        value_frame = ttk.LabelFrame(parent, text="条件值", padding=5)
        value_frame.pack(fill=tk.X)
        
        self.condition_value = ttk.Entry(value_frame)
        self.condition_value.pack(fill=tk.X, padx=5)
        
    def create_text_widgets(self, parent):
        """创建文本提取设置界面"""
        # 提取类型
        type_frame = ttk.LabelFrame(parent, text="提取类型", padding=5)
        type_frame.pack(fill=tk.X, pady=(0,5))
        
        self.text_type = tk.StringVar(value="left")
        ttk.Radiobutton(type_frame, text="左侧", value="left", 
                       variable=self.text_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="右侧", value="right", 
                       variable=self.text_type).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="中间", value="mid", 
                       variable=self.text_type).pack(side=tk.LEFT, padx=5)
        
        # 字符数量
        num_frame = ttk.LabelFrame(parent, text="字符数量", padding=5)
        num_frame.pack(fill=tk.X)
        
        self.char_num = ttk.Spinbox(num_frame, from_=1, to=100, width=10)
        self.char_num.set(1)
        self.char_num.pack(padx=5)

    def create_default_widgets(self, parent):
        """创建默认设置界面"""
        # 提示标签
        ttk.Label(parent, 
                 text="此公式暂无特殊设置选项",
                 font=(self.config.get('style.font_family'), 
                       self.config.get('style.font_size'))).pack(pady=20)

    def bind_events(self, callbacks):
        """绑定事件回调
        Args:
            callbacks: 回调函数字典
        """
        self.callbacks = callbacks
        
        # 按钮事件
        self.ok_btn.config(command=self.on_ok)
        self.cancel_btn.config(command=self.on_cancel)
        
        # 窗口关闭事件
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def get_settings(self):
        """获取设置值"""
        try:
            if self.formula_type == "条件求和":
                return {
                    'type': self.condition_type.get(),
                    'value': self.condition_value.get()
                }
            elif self.formula_type == "条件计数":
                return {
                    'type': self.count_type.get(),
                    'value': self.condition_value.get()
                }
            elif self.formula_type == "文本提取":
                return {
                    'type': self.text_type.get(),
                    'num': int(self.char_num.get())
                }
            else:
                return {}
            
        except Exception as e:
            print(f"获取设置值失败: {str(e)}")
            return {}

    def on_ok(self):
        """确定按钮事件处理"""
        try:
            # 获取设置值
            settings = self.get_settings()
            
            # 调用回调函数
            if 'on_ok' in self.callbacks:
                self.callbacks['on_ok'](settings)
            
            # 关闭窗口
            self.destroy()
            
        except Exception as e:
            print(f"确定按钮事件处理失败: {str(e)}")

    def on_cancel(self):
        """取消按钮事件处理"""
        try:
            # 调用回调函数
            if 'on_cancel' in self.callbacks:
                self.callbacks['on_cancel']()
            
            # 关闭窗口
            self.destroy()
            
        except Exception as e:
            print(f"取消按钮事件处理失败: {str(e)}") 