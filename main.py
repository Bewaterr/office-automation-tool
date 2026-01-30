from core.config import ConfigManager
from core.engine import OfficeEngine
from gui.main_window import MainWindow

def main():
    # 初始化配置管理器
    config = ConfigManager()
    
    # 初始化执行引擎
    engine = OfficeEngine(config)
    
    # 创建并运行主窗口
    window = MainWindow(config, engine)
    window.run()


if __name__ == "__main__":
    main()