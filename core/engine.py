import pythoncom
import win32com.client
from functools import wraps
import uuid


class OfficeEngine:
    """Office 自动化执行引擎"""
    
    def __init__(self, config_manager, logger=None):
        self.config = config_manager
        self.logger = logger
        self.active_functions = {}
    
    def log(self, message, level="info"):
        """记录日志"""
        if self.logger:
            self.logger(message, level)
        else:
            print(message)
    
    def create_handler(self, app_name, code, func_name):
        """动态创建函数处理器"""
        def handler():
            pythoncom.CoInitialize()
            try:
                # 获取 Office 应用
                app = win32com.client.GetActiveObject(f"{app_name}.Application")
                
                # 执行用户代码
                local_vars = {"app": app}
                exec(code, {}, local_vars)
                
                # 调用用户定义的函数
                if func_name in local_vars:
                    local_vars[func_name](app)
                    self.log(f"✅ {func_name} 执行成功")
                else:
                    self.log(f"❌ 找不到函数 {func_name}", "error")
                    
            except Exception as e:
                self.log(f"❌ {func_name} 失败: {e}", "error")
            finally:
                pythoncom.CoUninitialize()
        
        return handler
    
    def load_all_functions(self):
        """加载所有启用的功能"""
        self.active_functions.clear()
        
        for app_name, plugin in self.config.plugins.items():
            for func in plugin.get("functions", []):
                if func.get("enabled", True):
                    hotkey = func.get("hotkey")
                    code = func.get("code")
                    func_name = func.get("func_name")
                    
                    if hotkey and code and func_name:
                        handler = self.create_handler(app_name, code, func_name)
                        self.active_functions[hotkey] = handler
                        self.log(f"📌 加载: {hotkey} -> {func.get('name', func_name)}")
        
        return self.active_functions
    
    def check_hotkey_conflict(self, hotkey, exclude_id=None):
        """检查快捷键冲突"""
        for app_name, plugin in self.config.plugins.items():
            for func in plugin.get("functions", []):
                if func.get("id") != exclude_id and func.get("hotkey") == hotkey:
                    return func.get("name", "未命名功能")
        return None