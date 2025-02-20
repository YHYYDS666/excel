import tkinter as tk
from tkinter import ttk
from typing import Dict

class ApplyPage(ttk.Frame):
    """应用公式页面"""
    def __init__(self, parent, config, formula_manager, excel_manager):
        super().__init__(parent)
        self.parent = parent
        self.config = config
        self.formula_manager = formula_manager
        self.excel_manager = excel_manager
        
        # 初始化变量
        self.current_formula = None
        self.input_range = None
        self.output_range = None
        
        # 初始化事件回调
        self.callbacks = {}
        
        # 创建界面组件
        self.create_widgets()
    
    def create_widgets(self):
        """创建界面组件"""
        # 顶部工具栏
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, pady=(0,5))
        
        # 返回按钮
        self.back_btn = ttk.Button(
            toolbar,
            text="返回",
            width=8,
            command=lambda: self.callbacks['on_back']()
        )
        self.back_btn.pack(side=tk.LEFT, padx=2)
        
        # 应用按钮
        self.apply_btn = ttk.Button(
            toolbar,
            text="应用公式",
            width=10,
            command=self.apply_formula,
            state='disabled'  # 初始禁用
        )
        self.apply_btn.pack(side=tk.LEFT, padx=2)
        
        # 清除选择按钮
        self.clear_btn = ttk.Button(
            toolbar,
            text="清除选择",
            width=8,
            command=self.clear_selection
        )
        self.clear_btn.pack(side=tk.LEFT, padx=2)
        
        # 公式信息区域
        info_frame = ttk.LabelFrame(self, text="公式信息", padding=10)
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.formula_label = ttk.Label(info_frame, text="")
        self.formula_label.pack(fill=tk.X)
        
        self.desc_label = ttk.Label(info_frame, text="", wraplength=500)
        self.desc_label.pack(fill=tk.X)
        
        # 选择区域框架
        range_frame = ttk.LabelFrame(self, text="选择区域", padding=10)
        range_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 输入区域
        input_frame = ttk.Frame(range_frame)
        input_frame.pack(fill=tk.X, pady=(0,5))
        
        ttk.Label(input_frame, text="输入区域:").pack(side=tk.LEFT)
        
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(
            input_frame,
            textvariable=self.input_var,
            state='readonly'
        )
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.input_btn = ttk.Button(
            input_frame,
            text="选择",
            command=lambda: self.select_range('input'),
            style='Select.TButton'  # 使用绿色按钮样式
        )
        self.input_btn.pack(side=tk.LEFT)
        
        # 输出区域
        output_frame = ttk.Frame(range_frame)
        output_frame.pack(fill=tk.X)
        
        ttk.Label(output_frame, text="输出区域:").pack(side=tk.LEFT)
        
        self.output_var = tk.StringVar()
        self.output_entry = ttk.Entry(
            output_frame,
            textvariable=self.output_var,
            state='readonly'
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.output_btn = ttk.Button(
            output_frame,
            text="选择",
            command=lambda: self.select_range('output'),
            state='disabled',  # 初始禁用
            style='Select.TButton'  # 使用绿色按钮样式
        )
        self.output_btn.pack(side=tk.LEFT)
    
    def bind_events(self, callbacks):
        """绑定事件回调"""
        self.callbacks = callbacks
    
    def set_formula(self, formula_name):
        """设置当前公式"""
        self.current_formula = formula_name
        description = self.formula_manager.get_formula_description(formula_name)
        
        self.formula_label.config(text=f"公式: {formula_name}")
        self.desc_label.config(text=f"说明: {description}")
        
        # 清空选择区域
        self.input_var.set("")
        self.output_var.set("")
        self.input_range = None
        self.output_range = None
    
    def select_range(self, mode):
        """选择区域
        Args:
            mode: 选择模式('input'/'output')
        """
        try:
            # 保存当前活动窗口
            active_window = self.winfo_toplevel()
            
            # 最小化当前窗口以方便选择Excel区域
            active_window.iconify()
            
            # 等待一下以确保窗口最小化完成
            self.after(100)
            
            if 'on_range_selected' in self.callbacks:
                if self.callbacks['on_range_selected'](mode):
                    # 选择成功后恢复窗口
                    active_window.deiconify()
                    active_window.lift()
                    
                    # 更新状态栏
                    if 'on_status_update' in self.callbacks:
                        area_type = "输入" if mode == 'input' else "输出"
                        self.callbacks['on_status_update'](f"已选择{area_type}区域")
                else:
                    # 选择失败也要恢复窗口
                    active_window.deiconify()
                    active_window.lift()
                    
                    if 'on_status_update' in self.callbacks:
                        self.callbacks['on_status_update']("选择区域失败")
                    
        except Exception as e:
            print(f"选择区域失败: {str(e)}")
            # 确保窗口恢复
            active_window.deiconify()
            active_window.lift()
    
    def update_range_display(self):
        """更新区域显示"""
        if self.input_range:
            self.input_var.set(self.input_range.Address)
        if self.output_range:
            self.output_var.set(self.output_range.Address)
    
    def update_selection(self, range_obj, mode='input'):
        """更新选择区域
        Args:
            range_obj: 选择的区域对象
            mode: 选择模式('input'/'output')
        """
        try:
            if mode == 'input':
                self.input_range = range_obj
                self.input_var.set(range_obj.Address)
                # 启用输出区域选择按钮，但不禁用输入区域选择
                self.output_btn['state'] = 'normal'
            else:
                self.output_range = range_obj
                self.output_var.set(range_obj.Address)
            
            # 如果已选择输入区域，启用应用按钮
            if self.input_range:
                self.apply_btn['state'] = 'normal'
            
        except Exception as e:
            print(f"更新选择区域失败: {str(e)}")
    
    def clear_selection(self):
        """清空选择"""
        self.input_range = None
        self.output_range = None
        self.input_var.set('')
        self.output_var.set('')
        self.output_btn['state'] = 'disabled'
        self.apply_btn['state'] = 'disabled'
    
    def apply_formula(self):
        """应用公式"""
        if not self.input_range:
            if 'on_status_update' in self.callbacks:
                self.callbacks['on_status_update']("请先选择输入区域")
            return
            
        if 'on_apply' in self.callbacks:
            self.callbacks['on_apply'](
                self.current_formula,
                self.input_range,
                self.output_range
            )
    
    def update_formula_info(self, formula: Dict):
        """更新公式信息
        Args:
            formula: 公式信息字典
        """
        try:
            self.current_formula = formula
            self.formula_label.config(text=f"公式: {formula['name']}")
            self.desc_label.config(text=f"说明: {formula['description']}")
            
            # 清空选择区域
            self.input_var.set('')
            self.output_var.set('')
            self.input_range = None
            self.output_range = None
            
        except Exception as e:
            print(f"更新公式信息失败: {str(e)}")

    def show(self):
        """显示页面"""
        self.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def hide(self):
        """隐藏页面"""
        self.pack_forget() 