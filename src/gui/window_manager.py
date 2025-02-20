class WindowManager:
    """窗口管理基类"""
    def __init__(self, root, config):
        self.root = root
        self.config = config
        
        # 绑定窗口事件
        self.bind_window_events()
        
        # 恢复窗口状态
        self.restore_window_state()
        
    def setup_window(self):
        """设置窗口属性"""
        # 获取配置
        window_width = self.config.get('window.width', 1000)
        window_height = self.config.get('window.height', 600)
        min_width = self.config.get('window.min_width', 800)
        min_height = self.config.get('window.min_height', 500)
        
        # 计算窗口位置
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # 设置窗口属性
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(min_width, min_height)
        self.root.resizable(True, True)
        
        # 设置窗口置顶(仅在打开时)
        self.root.attributes('-topmost', True)
        self.root.after(1000, lambda: self.root.attributes('-topmost', False)) 

    def save_window_state(self):
        """保存窗口状态"""
        try:
            # 保存窗口位置和大小
            self.config.set('window.x', self.root.winfo_x())
            self.config.set('window.y', self.root.winfo_y())
            self.config.set('window.width', self.root.winfo_width())
            self.config.set('window.height', self.root.winfo_height())
            
            # 保存窗口是否最大化
            self.config.set('window.zoomed', self.root.state() == 'zoomed')
            
        except Exception as e:
            print(f"保存窗口状态失败: {str(e)}")

    def restore_window_state(self):
        """恢复窗口状态"""
        try:
            # 恢复窗口位置和大小
            x = self.config.get('window.x')
            y = self.config.get('window.y')
            width = self.config.get('window.width', 1000)
            height = self.config.get('window.height', 600)
            
            if all(v is not None for v in (x, y)):
                self.root.geometry(f"{width}x{height}+{x}+{y}")
            
            # 恢复最大化状态
            if self.config.get('window.zoomed', False):
                self.root.state('zoomed')
            
        except Exception as e:
            print(f"恢复窗口状态失败: {str(e)}")

    def bind_window_events(self):
        """绑定窗口事件"""
        try:
            # 窗口大小改变事件
            self.root.bind('<Configure>', self.on_window_configure)
            
            # 窗口关闭事件
            self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)
            
        except Exception as e:
            print(f"绑定窗口事件失败: {str(e)}")

    def on_window_configure(self, event):
        """窗口大小改变事件处理"""
        try:
            # 忽略非窗口的Configure事件
            if event.widget != self.root:
                return
            
            # 保存窗口状态
            self.save_window_state()
            
        except Exception as e:
            print(f"处理窗口大小改变事件失败: {str(e)}")

    def on_window_close(self):
        """窗口关闭事件处理"""
        try:
            # 保存窗口状态
            self.save_window_state()
            
            # 调用子类的关闭处理
            if hasattr(self, 'on_closing'):
                self.on_closing()
            else:
                self.root.destroy()
            
        except Exception as e:
            print(f"处理窗口关闭事件失败: {str(e)}")
            self.root.destroy() 