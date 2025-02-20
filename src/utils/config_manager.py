import os
import json
from typing import Any, Dict, Optional

class ConfigManager:
    """配置管理器"""
    
    def __init__(self, config_file="config.json"):
        self.config_file = config_file
        self.config = self.load_config()
        
    def load_config(self):
        """加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"加载配置失败: {str(e)}")
            return {}
            
    def save_config(self):
        """保存配置"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"保存配置失败: {str(e)}")
            
    def get(self, key, default=None):
        """获取配置项"""
        return self.config.get(key, default)
            
    def set(self, key, value):
        """设置配置项"""
        self.config[key] = value
        self.save_config()
        
    def merge_config(self, default: Dict, user: Dict) -> Dict:
        """合并配置
        Args:
            default: 默认配置
            user: 用户配置
        Returns:
            Dict: 合并后的配置
        """
        result = default.copy()
        
        for key, value in user.items():
            if key in result and isinstance(value, dict):
                result[key] = self.merge_config(result[key], value)
            else:
                result[key] = value
                
        return result 

    def save(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"保存配置失败: {str(e)}")
            return False 

    def load_custom_config(self):
        """加载自定义配置"""
        try:
            if os.path.exists('custom_config.json'):
                with open('custom_config.json', 'r', encoding='utf-8') as f:
                    custom_config = json.load(f)
                    self.config.update(custom_config)
        except Exception as e:
            print(f"加载自定义配置失败: {str(e)}")

    def save_custom_config(self):
        """保存自定义配置"""
        try:
            # 只保存需要持久化的配置项
            custom_config = {
                'window.width': self.config.get('window.width'),
                'window.height': self.config.get('window.height'),
                'formula.auto_expand': self.config.get('formula.auto_expand'),
                'formula.search_delay': self.config.get('formula.search_delay'),
                'recent_files': self.config.get('recent_files', [])
            }
            
            with open('custom_config.json', 'w', encoding='utf-8') as f:
                json.dump(custom_config, f, indent=4, ensure_ascii=False)
            
        except Exception as e:
            print(f"保存自定义配置失败: {str(e)}") 