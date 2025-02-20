import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox
from gui.main_window import MainWindow
from utils.config_manager import ConfigManager
import logging
from logging.handlers import RotatingFileHandler
import argparse

VERSION = '1.0.0'

def check_dependencies():
    """检查依赖项"""
    try:
        import win32com.client
        return True
    except ImportError:
        messagebox.showerror(
            "错误",
            "缺少必要的依赖项。\n请安装 pywin32:\npip install pywin32"
        )
        return False

def setup_environment():
    """设置运行环境"""
    try:
        # 获取当前文件路径
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 将src目录添加到系统路径
        if current_dir not in sys.path:
            sys.path.append(current_dir)
            
        # 获取项目根目录
        root_dir = os.path.dirname(current_dir)
        
        # 确保data目录存在
        data_dir = os.path.join(root_dir, "data")
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
            
        # 设置DPI感知
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
            
        return True
        
    except Exception as e:
        print(f"设置运行环境失败: {str(e)}")
        return False

def setup_style(config):
    """设置应用程序样式"""
    try:
        style = ttk.Style()
        
        # 获取样式配置
        font_family = config.get('style.font_family', '微软雅黑')
        font_size = config.get('style.font_size', 10)
        colors = config.get('style.colors', {
            'background': '#ffffff',
            'foreground': '#333333',
            'header': '#f0f0f0',
            'button': '#e1e1e1',
            'highlight': '#0078d7'
        })
        
        # 配置通用样式
        style.configure(
            ".",
            font=(font_family, font_size),
            background=colors['background']
        )
        
        # 配置Frame样式
        style.configure(
            "TFrame",
            background=colors['background']
        )
        
        # 配置Label样式
        style.configure(
            "TLabel",
            background=colors['background'],
            foreground=colors['foreground']
        )
        
        # 配置Button样式
        style.configure(
            "TButton",
            padding=5,
            background=colors['button']
        )
        
        # 配置Treeview样式
        style.configure(
            "Treeview",
            background=colors['background'],
            fieldbackground=colors['background'],
            foreground=colors['foreground'],
            rowheight=25
        )
        
        style.configure(
            "Treeview.Heading",
            background=colors['header'],
            foreground=colors['foreground'],
            padding=5
        )
        
        return True
        
    except Exception as e:
        print(f"设置样式失败: {str(e)}")
        return False

def setup_logging(level='INFO'):
    """设置日志"""
    # 创建日志目录
    log_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    # 设置日志格式
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 文件处理器(限制大小和数量)
    file_handler = RotatingFileHandler(
        os.path.join(log_dir, 'app.log'),
        maxBytes=1024*1024,  # 1MB
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setFormatter(formatter)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    # 配置根日志器
    logging.basicConfig(
        level=level,
        handlers=[file_handler, console_handler]
    )

def setup_exception_handler():
    """设置全局异常处理"""
    def handle_exception(exc_type, exc_value, exc_traceback):
        # 记录异常
        logging.error(
            "Uncaught exception",
            exc_info=(exc_type, exc_value, exc_traceback)
        )
        # 显示错误消息
        messagebox.showerror(
            "错误",
            f"发生未处理的异常:\n{str(exc_value)}"
        )
    
    # 设置异常处理器
    sys.excepthook = handle_exception

def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description='Excel公式工具 v' + VERSION,
        epilog='使用 --help 查看更多帮助信息'
    )
    
    parser.add_argument(
        '-v', '--version',
        action='version',
        version=f'%(prog)s {VERSION}'
    )
    
    parser.add_argument(
        '--config',
        help='配置文件路径',
        default='config.json'
    )
    
    parser.add_argument(
        '--debug',
        action='store_true',
        help='启用调试模式'
    )
    
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='日志级别'
    )
    
    return parser.parse_args()

def main():
    """主函数"""
    try:
        # 解析命令行参数
        args = parse_args()
        
        # 设置日志级别
        setup_logging(level=args.log_level)
        
        # 启用调试模式
        if args.debug:
            logging.getLogger().setLevel(logging.DEBUG)
            
        setup_exception_handler()
        logging.info("程序启动")
        
        # 检查依赖项
        if not check_dependencies():
            return
            
        # 设置运行环境
        if not setup_environment():
            messagebox.showerror("错误", "初始化环境失败")
            return
            
        # 加载配置
        config = ConfigManager(args.config)
        
        # 创建主窗口
        root = tk.Tk()
        root.title("Excel公式工具")
        
        # 设置窗口图标
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "..", "assets", "icon.ico")
            if os.path.exists(icon_path):
                root.iconbitmap(icon_path)
        except:
            pass
            
        # 设置窗口大小和位置
        window_width = config.get('window.width', 1000)
        window_height = config.get('window.height', 600)
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置最小窗口大小
        min_width = config.get('window.min_width', 800)
        min_height = config.get('window.min_height', 500)
        root.minsize(min_width, min_height)
        
        # 设置应用程序样式
        setup_style(config)
        
        # 创建主应用
        app = MainWindow(root, config)
        
        # 启动主循环
        root.mainloop()
        
    except Exception as e:
        logging.error(f"程序运行失败: {str(e)}", exc_info=True)
        messagebox.showerror("错误", f"程序运行失败: {str(e)}")

if __name__ == "__main__":
    main() 