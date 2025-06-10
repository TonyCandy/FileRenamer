自然语言开发全过程https://www.doubao.com/thread/w0e1af4e11cccdfdd
有以下功能：
1、可以批量重命名指定路径下的子文件夹
2、可以批量重命名指定路径下的所有文件
使用pyinstaller --onefile --windowed main.py打包成exe文件
Python 3.13
安装 PyQt5 和 pandas
pip install PyQt5 pandas
安装 PyInstaller
pip install pyinstaller
在项目文件夹中执行：pyinstaller --onefile --windowed main.py

### 一、GitHub 项目准备

1. **项目名称**: FileRenamer
2. **README.md 内容结构**:
```markdown
# FileRenamer - 文件批量重命名工具

一个基于 PyQt5 的现代化文件批量重命名工具，支持文件和文件夹的灵活重命名操作。

## 特性
- 支持文件和文件夹的批量重命名
- 支持拖拽操作
- Excel 导出与导入功能
- 自然排序算法（与 Windows 资源管理器一致）
- 详细的文件信息显示
- 递归获取子文件夹内容
- 实时状态更新
- 友好的用户界面

## 安装
```bash
# 安装依赖
pip install PyQt5 pandas pillow
```

## 使用方法
1. 运行程序
2. 选择或拖拽文件/文件夹
3. 在表格中编辑新名称
4. 点击"执行重命名"完成操作

## 开发环境
- Python 3.8+
- PyQt5
- pandas
- pillow

## 许可证
MIT License
```
