from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QScrollArea,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QHBoxLayout,
    QSplitter,
    QTextEdit,
    QMenuBar,
    QFileDialog,
)
from docx import Document
import re
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QTextOption


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # 设置窗口标题和初始大小
        self.setWindowTitle("Automatic Document Filler")
        self.setGeometry(100, 100, 1600, 900)
        self.init_ui()
        self.optimize_update()

    def init_ui(self):
        # 添加菜单栏
        menu_bar = QMenuBar(self)
        file_menu = menu_bar.addMenu("文件")
        edit_menu = menu_bar.addMenu("报价")
        settings_menu = menu_bar.addMenu("帮助")
        self.setMenuBar(menu_bar)

        # 创建主布局
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)

        # 左侧固定宽度区域（包含表单和提交按钮）
        left_widget = QWidget()
        left_widget.setFixedWidth(500)  # 固定宽度为 500
        left_layout = QVBoxLayout(left_widget)

        # 创建滚动视图
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        form_container = QWidget()
        form_layout = QVBoxLayout(form_container)

        # 添加表单内容到滚动视图
        self.inputs = {}
        fields = [
            "Company Name",
            "Address",
            "Contact Person",
            "Phone Number",
            "Email",
            "Website",
            "Department",
            "Department",
            "Department",
            "Department",
            "Department",
            "Department",
        ]
        font_label = QFont("Segoe UI", 10)
        font_input = QFont("Segoe UI", 10)

        form_layout.setSpacing(15)  # 设置表单项之间的垂直间隔
        form_layout.setContentsMargins(10, 10, 10, 10)  # 设置表单内容边距

        for field in fields:
            label = QLabel(field)
            label.setFont(font_label)
            input_box = QLineEdit()
            input_box.setFont(font_input)
            input_box.setStyleSheet(
                "QLineEdit {"
                "padding: 8px;"
                "border: 1px solid #ccc;"
                "border-radius: 4px;"
                "background-color: #f9f9f9;"
                "}"
                "QLineEdit:focus {"
                "border: 1px solid #007BFF;"
                "background-color: #fff;"
                "}"
            )
            input_box.textChanged.connect(self.update_preview)  # 实时更新预览
            self.inputs[field] = input_box
            form_layout.addWidget(label)
            form_layout.addWidget(input_box)

        form_container.setLayout(form_layout)
        scroll_area.setWidget(form_container)

        # 将滚动视图添加到左侧布局
        left_layout.addWidget(scroll_area)

        # 添加提交按钮
        generate_button = QPushButton("Generate Document")
        generate_button.setFont(font_label)
        generate_button.setStyleSheet(
            "QPushButton {"
            "padding: 10px;"
            "background-color: #007BFF;"
            "color: white;"
            "border: none;"
            "border-radius: 4px;"
            "}"
            "QPushButton:hover {"
            "background-color: #0056b3;"
            "}"
        )
        generate_button.clicked.connect(self.generate_document)
        left_layout.addWidget(generate_button, alignment=Qt.AlignBottom)

        # 右侧预览区域
        self.preview = QTextEdit()
        self.preview.setFont(QFont("Segoe UI", 10))
        self.preview.setReadOnly(True)
        self.preview.setStyleSheet(
            "QTextEdit {"
            "padding: 10px;"
            "border: 1px solid #ccc;"
            "border-radius: 4px;"
            "background-color: #f5f5f5;"
            "}"
        )
        self.preview.setWordWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)

        # 将左侧和右侧添加到主布局
        main_layout.addWidget(left_widget)
        main_layout.addWidget(self.preview)

        # 设置主窗口布局
        self.setCentralWidget(main_widget)

    def optimize_update(self):
        self.update_timer = QTimer(self)
        self.update_timer.setSingleShot(True)
        self.update_timer.timeout.connect(self.refresh_preview)

    def update_preview(self):
        # 在输入时启动延迟更新
        self.update_timer.start(300)  # 延迟 300 毫秒

    def refresh_preview(self):
       # 实时更新预览内容
        preview_text = "<h3>Preview:</h3>"
        for field, input_box in self.inputs.items():
            preview_text += f"<p><b>{field}:</b> {input_box.text()}</p>"
        self.preview.setHtml(preview_text)

    def generate_document(self):
        # 实现生成文档的逻辑
        print("Generate document button clicked")

    def load_template(self):
        """加载 Word 模板文件"""
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Word Template", "", "Word Files (*.docx)")
        if file_path:
            self.template_document = Document(file_path)
            self.template_content = self.word_to_html(self.template_document)
            self.preview.setHtml(self.template_content)  # 显示模板内容


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()
