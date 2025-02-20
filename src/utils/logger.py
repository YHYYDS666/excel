import logging
import os
from datetime import datetime

def setup_logger():
    """配置日志记录器"""
    # 创建logs目录
    if not os.path.exists('logs'):
        os.makedirs('logs')
        
    # 配置日志格式
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.FileHandler(f'logs/app_{datetime.now().strftime("%Y%m%d")}.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger('ExcelTool') 