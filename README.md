# Office 自动化插件管理器

一个强大的 Office 自动化工具，支持 PowerPoint、Excel、Word。

## 安装

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 运行程序：
```bash
python main.py
```

## 打包成 EXE

```bash
pip install pyinstaller
pyinstaller --name="OfficeAutomationTool" --windowed --onefile main.py
```

## 功能特点

- ✅ 可视化管理快捷键功能
- ✅ 在线编辑/添加新功能
- ✅ 功能分组管理
- ✅ 启用/禁用功能
- ✅ 快捷键冲突检测
- ✅ 导入/导出配置
- ✅ 实时日志查看

## 使用说明

1. 启动程序后，左侧显示所有功能列表
2. 点击功能可在右侧编辑代码
3. 点击"新建"添加新功能
4. 点击"保存"保存修改
5. 点击"测试运行"验证代码

## 快捷键格式

- 单键: `<alt>+a`, `<ctrl>+b`
- 组合键: `<ctrl>+<shift>+c`
