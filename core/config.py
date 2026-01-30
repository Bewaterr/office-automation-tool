import json
import os
from pathlib import Path


class ConfigManager:
    """配置管理器"""
    
    def __init__(self, config_dir="plugins"):
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(exist_ok=True)
        self.plugins = {}
        self.load_all_plugins()
    
    def load_all_plugins(self):
        """加载所有插件配置"""
        for file in self.config_dir.glob("*.json"):
            app_name = file.stem
            with open(file, 'r', encoding='utf-8') as f:
                self.plugins[app_name] = json.load(f)
    
    def save_plugin(self, app_name, plugin_data):
        """保存插件配置"""
        file_path = self.config_dir / f"{app_name}.json"
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(plugin_data, f, indent=2, ensure_ascii=False)
        self.plugins[app_name] = plugin_data
    
    def get_plugin(self, app_name):
        """获取插件配置"""
        return self.plugins.get(app_name, {"functions": []})
    
    def add_function(self, app_name, func_data):
        """添加新功能"""
        if app_name not in self.plugins:
            self.plugins[app_name] = {"functions": []}
        
        self.plugins[app_name]["functions"].append(func_data)
        self.save_plugin(app_name, self.plugins[app_name])
    
    def update_function(self, app_name, func_id, func_data):
        """更新功能"""
        plugin = self.plugins.get(app_name)
        if plugin:
            for i, func in enumerate(plugin["functions"]):
                if func.get("id") == func_id:
                    plugin["functions"][i] = func_data
                    self.save_plugin(app_name, plugin)
                    return True
        return False
    
    def delete_function(self, app_name, func_id):
        """删除功能"""
        plugin = self.plugins.get(app_name)
        if plugin:
            plugin["functions"] = [f for f in plugin["functions"] if f.get("id") != func_id]
            self.save_plugin(app_name, plugin)
            return True
        return False
    
    def export_config(self, filepath):
        """导出所有配置"""
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(self.plugins, f, indent=2, ensure_ascii=False)
    
    def import_config(self, filepath):
        """导入配置"""
        with open(filepath, 'r', encoding='utf-8') as f:
            imported = json.load(f)
            for app_name, plugin_data in imported.items():
                self.save_plugin(app_name, plugin_data)