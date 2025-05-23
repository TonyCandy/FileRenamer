import sys
import os
import pandas as pd
import tempfile
import subprocess
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QLineEdit, 
                            QTableWidget, QTableWidgetItem, QFileDialog, 
                            QMessageBox, QHeaderView, QDialog, QRadioButton,
                            QButtonGroup, QDialogButtonBox)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QColor, QDragEnterEvent, QDropEvent

class DragDropDialog(QDialog):
    """拖拽文件/文件夹后的选择对话框"""
    def __init__(self, paths, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择操作")
        self.paths = paths
        
        layout = QVBoxLayout(self)
        
        # 根据拖放内容显示不同的提示
        if len(paths) == 1:
            item_type = "文件夹" if os.path.isdir(paths[0]) else "文件"
            layout.addWidget(QLabel(f"已拖入1个{item_type}，请选择操作:"))
        else:
            layout.addWidget(QLabel(f"已拖入{len(paths)}个项目，请选择操作:"))
        
        # 操作选项
        self.type_group = QButtonGroup(self)
        
        self.file_radio = QRadioButton("获取所有文件")
        self.folder_radio = QRadioButton("获取所有文件夹")
        
        self.type_group.addButton(self.file_radio, 0)
        self.type_group.addButton(self.folder_radio, 1)
        
        # 默认选择与拖放内容匹配的选项
        if all(os.path.isdir(p) for p in paths):
            self.folder_radio.setChecked(True)
        else:
            self.file_radio.setChecked(True)
        
        layout.addWidget(self.file_radio)
        layout.addWidget(self.folder_radio)
        
        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
    def get_selected_option(self):
        """获取用户选择的操作类型"""
        if self.file_radio.isChecked():
            return "file"
        return "folder"

class FileRenamer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文件批量重命名工具")
        self.setGeometry(100, 100, 1000, 600)
        
        # 启用拖拽功能
        self.setAcceptDrops(True)
        
        # 创建中心部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 路径选择部分
        path_layout = QHBoxLayout()
        self.path_label = QLabel("文件夹路径:")
        self.path_input = QLineEdit()
        self.browse_button = QPushButton("浏览...")
        self.browse_button.clicked.connect(self.browse_path)
        
        path_layout.addWidget(self.path_label)
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.browse_button)
        main_layout.addLayout(path_layout)
        
        # 获取文件/文件夹按钮
        get_layout = QHBoxLayout()
        self.get_files_button = QPushButton("获取文件")
        self.get_files_button.clicked.connect(lambda: self.get_items(item_type="file"))
        
        self.get_folders_button = QPushButton("获取文件夹")
        self.get_folders_button.clicked.connect(lambda: self.get_items(item_type="folder"))
        
        self.clear_button = QPushButton("清除信息")
        self.clear_button.clicked.connect(self.clear_items)
        
        get_layout.addWidget(self.get_files_button)
        get_layout.addWidget(self.get_folders_button)
        get_layout.addWidget(self.clear_button)
        main_layout.addLayout(get_layout)
        
        # 表格部分 - 增加了路径和状态列
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["类型", "名称", "扩展名", "新名称", "路径", "状态"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        main_layout.addWidget(self.table)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        self.open_excel_button = QPushButton("打开重命名表格")
        self.open_excel_button.clicked.connect(self.open_excel)
        
        self.refresh_button = QPushButton("刷新表格")
        self.refresh_button.clicked.connect(self.refresh_table)
        
        self.rename_button = QPushButton("执行重命名")
        self.rename_button.clicked.connect(self.rename_files)
        
        button_layout.addWidget(self.open_excel_button)
        button_layout.addWidget(self.refresh_button)
        button_layout.addWidget(self.rename_button)
        main_layout.addLayout(button_layout)
        
        # 状态标签
        self.status_label = QLabel("就绪")
        main_layout.addWidget(self.status_label)
        
        # 数据存储
        self.item_list = []  # 存储所有文件/文件夹信息
        self.temp_excel_file = os.path.join(tempfile.gettempdir(), "file_renamer_temp.xlsx")
        
    def browse_path(self):
        """浏览并选择文件夹路径"""
        folder_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder_path:
            self.path_input.setText(folder_path)
    
    def get_items(self, item_type, paths=None):
        """获取指定路径下的文件或文件夹"""
        # 如果没有指定路径，使用输入框中的路径
        if paths is None:
            paths = [self.path_input.text()]
        
        total_added = 0
        
        for path in paths:
            if not os.path.exists(path):
                QMessageBox.warning(self, "错误", f"指定的路径不存在: {path}")
                continue
            
            try:
                count_before = len(self.item_list)
                
                # 获取文件或文件夹列表
                if item_type == "file":
                    # 获取文件
                    if os.path.isdir(path):
                        for item in os.listdir(path):
                            item_path = os.path.join(path, item)
                            if os.path.isfile(item_path):
                                name, ext = os.path.splitext(item)
                                self.item_list.append({
                                    "type": "文件",
                                    "name": name,
                                    "ext": ext,
                                    "new_name": name,
                                    "path": path,
                                    "status": ""
                                })
                    else:
                        # 如果是单个文件，直接添加
                        name, ext = os.path.splitext(os.path.basename(path))
                        self.item_list.append({
                            "type": "文件",
                            "name": name,
                            "ext": ext,
                            "new_name": name,
                            "path": os.path.dirname(path),
                            "status": ""
                        })
                else:
                    # 获取文件夹
                    if os.path.isdir(path):
                        # 添加当前文件夹
                        self.item_list.append({
                            "type": "文件夹",
                            "name": os.path.basename(path),
                            "ext": "文件夹",
                            "new_name": os.path.basename(path),
                            "path": os.path.dirname(path) if os.path.dirname(path) else path,
                            "status": ""
                        })
                        
                        # 添加子文件夹
                        for item in os.listdir(path):
                            item_path = os.path.join(path, item)
                            if os.path.isdir(item_path):
                                self.item_list.append({
                                    "type": "文件夹",
                                    "name": item,
                                    "ext": "文件夹",
                                    "new_name": item,
                                    "path": path,
                                    "status": ""
                                })
                
                # 更新表格
                count_added = len(self.item_list) - count_before
                total_added += count_added
                
            except Exception as e:
                QMessageBox.warning(self, "错误", f"获取项目列表时出错: {str(e)}")
        
        if total_added > 0:
            self.update_table()
            self.status_label.setText(f"已添加 {total_added} 个项目，共 {len(self.item_list)} 个项目")
    
    def clear_items(self):
        """清除所有项目信息"""
        self.item_list = []
        self.update_table()
        self.status_label.setText("已清除所有信息")
    
    def update_table(self):
        """更新表格显示"""
        self.table.setRowCount(len(self.item_list))
        for row, item in enumerate(self.item_list):
            # 设置类型列
            self.table.setItem(row, 0, QTableWidgetItem(item["type"]))
            
            # 设置名称列
            name_item = QTableWidgetItem(item["name"])
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 1, name_item)
            
            # 设置扩展名列
            ext_item = QTableWidgetItem(item["ext"])
            ext_item.setFlags(ext_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 2, ext_item)
            
            # 设置新名称列
            self.table.setItem(row, 3, QTableWidgetItem(item["new_name"]))
            
            # 设置路径列
            path_item = QTableWidgetItem(item["path"])
            path_item.setFlags(path_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 4, path_item)
            
            # 设置状态列，根据状态设置颜色
            status_item = QTableWidgetItem(str(item["status"]))  # 将状态转换为字符串
            if item["status"] == "已重命名":
                status_item.setBackground(QColor(144, 238, 144))  # 浅绿色
            elif item["status"] == "重命名失败":
                status_item.setBackground(QColor(255, 182, 193))  # 浅红色
            self.table.setItem(row, 5, status_item)
    
    def open_excel(self):
        """将数据导出到Excel并打开"""
        if not self.item_list:
            QMessageBox.information(self, "提示", "没有数据可导出!")
            return
        
        try:
            # 创建DataFrame并保存到临时Excel文件
            df = pd.DataFrame(self.item_list)
            df.to_excel(self.temp_excel_file, index=False)
            
            # 尝试打开Excel文件 - 使用更可靠的方法
            try:
                # 对于Windows系统
                os.startfile(self.temp_excel_file)
            except AttributeError:
                # 对于非Windows系统 (这里不会执行，但保留代码以防万一)
                subprocess.call(['open', self.temp_excel_file])
                
            self.status_label.setText(f"已打开Excel文件: {self.temp_excel_file}")
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"打开Excel文件时出错: {str(e)}")
    
    def refresh_table(self):
        """从Excel文件刷新表格数据"""
        if not os.path.exists(self.temp_excel_file):
            QMessageBox.warning(self, "错误", "Excel文件不存在，请先打开重命名表格!")
            return
        
        try:
            # 从Excel文件读取数据
            df = pd.read_excel(self.temp_excel_file)
            
            # 检查数据格式是否正确
            required_columns = {'type', 'name', 'ext', 'new_name', 'path', 'status'}
            if not required_columns.issubset(set(df.columns)):
                missing = required_columns - set(df.columns)
                QMessageBox.warning(self, "错误", f"Excel文件缺少必要的列: {', '.join(missing)}")
                return
                
            # 确保所有列都存在
            for col in required_columns:
                if col not in df.columns:
                    df[col] = ''
            
            # 对数据类型进行转换
            df['type'] = df['type'].astype(str)
            df['name'] = df['name'].astype(str)
            df['ext'] = df['ext'].astype(str)
            df['new_name'] = df['new_name'].astype(str)
            df['path'] = df['path'].astype(str)
            df['status'] = df['status'].astype(str)
            
            # 更新文件列表
            self.item_list = df.to_dict('records')
            
            # 更新表格
            self.update_table()
            self.status_label.setText("已从Excel文件刷新数据")
            
        except Exception as e:
            # 捕获更具体的错误信息
            import traceback
            error_msg = traceback.format_exc()
            QMessageBox.warning(self, "错误", f"读取Excel文件时出错:\n{str(e)}\n\n详细信息:\n{error_msg}")
    
    def rename_files(self):
        """执行重命名操作"""
        try:
            if not self.item_list:
                QMessageBox.information(self, "提示", "没有项目可重命名!")
                return
            
            # 确认对话框
            reply = QMessageBox.question(
                self, "确认重命名", 
                f"确定要重命名 {len(self.item_list)} 个项目吗?\n此操作不可撤销!",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                renamed_count = 0
                failed_count = 0
                
                for i, item in enumerate(self.item_list):
                    # 跳过不需要重命名的项目
                    if item["name"] + item["ext"] == item["new_name"] + (item["ext"] if item["type"] == "文件" else ""):
                        self.item_list[i]["status"] = "未修改"
                        continue
                        
                    old_name = item["name"] + item["ext"] if item["type"] == "文件" else item["name"]
                    new_name = item["new_name"] + item["ext"] if item["type"] == "文件" else item["new_name"]
                    
                    old_path = os.path.join(item["path"], old_name)
                    new_path = os.path.join(item["path"], new_name)
                    
                    try:
                        # 检查文件是否存在
                        if not os.path.exists(old_path):
                            self.item_list[i]["status"] = "文件不存在"
                            failed_count += 1
                            continue
                        
                        # 检查新名称是否已存在
                        if os.path.exists(new_path):
                            self.item_list[i]["status"] = "新名称已存在"
                            failed_count += 1
                            continue
                        
                        # 特殊字符检查
                        invalid_chars = r'<>:"/\|?*'
                        if any(c in new_name for c in invalid_chars):
                            self.item_list[i]["status"] = f"包含无效字符: {invalid_chars}"
                            failed_count += 1
                            continue
                        
                        # 执行重命名
                        os.rename(old_path, new_path)
                        self.item_list[i]["status"] = "已重命名"
                        
                        # 更新名称（以便可以连续重命名）
                        self.item_list[i]["name"] = item["new_name"]
                        renamed_count += 1
                    except Exception as e:
                        self.item_list[i]["status"] = f"失败: {str(e)}"
                        failed_count += 1
                
                # 更新表格
                self.update_table()
                
                # 保存结果到Excel
                try:
                    df = pd.DataFrame(self.item_list)
                    df.to_excel(self.temp_excel_file, index=False)
                except:
                    pass  # 忽略保存失败
                
                # 显示结果
                if failed_count == 0:
                    QMessageBox.information(self, "成功", f"成功重命名 {renamed_count} 个项目!")
                else:
                    message = f"成功重命名 {renamed_count} 个项目，失败 {failed_count} 个。"
                    QMessageBox.warning(self, "部分失败", message)
                
                self.status_label.setText(f"重命名完成: {renamed_count} 成功, {failed_count} 失败")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"重命名过程中出现错误: {str(e)}")
    
    # 拖拽相关方法
    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        """拖拽放下事件"""
        urls = event.mimeData().urls()
        paths = [url.toLocalFile() for url in urls]
        
        if not paths:
            return
        
        # 显示选择对话框
        dialog = DragDropDialog(paths, self)
        if dialog.exec_():
            action = dialog.get_selected_option()
            self.get_items(action, paths)
            
            # 如果只拖入了一个项目，更新路径输入框
            if len(paths) == 1:
                self.path_input.setText(paths[0])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileRenamer()
    window.show()
    sys.exit(app.exec_())