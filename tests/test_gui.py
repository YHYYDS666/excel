import unittest
import tkinter as tk
from src.gui.main_window import MainWindow

class TestGUI(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """测试前准备"""
        cls.root = tk.Tk()
        cls.app = MainWindow(cls.root, {})
        
    def test_window(self):
        """测试窗口创建"""
        self.assertIsInstance(self.app, MainWindow)
        
    def test_toolbar(self):
        """测试工具栏"""
        self.assertIsNotNone(self.app.toolbar_manager)
        
    def test_status(self):
        """测试状态栏"""
        self.assertIsNotNone(self.app.status_manager)
        
    @classmethod
    def tearDownClass(cls):
        """测试后清理"""
        cls.root.destroy()

if __name__ == '__main__':
    unittest.main() 