
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QSpinBox, QPushButton, QLineEdit
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
class CustomIntInputDialog(QDialog):
    def __init__(self, title, prompt, min_value, max_value, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(400)

        # 设置窗口样式
        self.setStyleSheet(
            """
            QDialog {
                background-color: #f5f5f5;
                border-radius: 8px;
                font-family: 'Segoe UI';
                font-size: 14px;
            }
            QLabel {
                font-size: 14px;
                color: #333;
                margin-bottom: 10px;
                font-weight: normal;  /* 设置字体粗细 */
            }
            QSpinBox {
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: white;
                font-size: 14px;
                font-family: 'Segoe UI';
            }
            QSpinBox::up-button, QSpinBox::down-button {
                subcontrol-origin: border;
                subcontrol-position: right;
                width: 24px;
                border-left: 1px solid #ccc;
                border-radius: 4px;
                background-color: #f9f9f9;
            }
            QSpinBox::up-button {
                subcontrol-position: top right;
                border-bottom: 1px solid #ccc;
                border-top-right-radius: 4px;
            }
            QSpinBox::down-button {
                subcontrol-position: bottom right;
                border-bottom-right-radius: 4px;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background-color: #e0e0e0;
            }
            QSpinBox::up-arrow {
                image: url(./icon/tianjia.png);  /* 如果需要自定义箭头图标，可以替换为实际路径 */
            }
            QSpinBox::down-arrow {
                image: url(./icon/zhedie.png);  /* 如果需要自定义箭头图标，可以替换为实际路径 */
            }
            QSpinBox::up-arrow, QSpinBox::down-arrow {
                width: 16px;  /* 调整宽度 */
                height: 16px; /* 调整高度 */
            }
            QPushButton {
                padding: 10px;
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                margin-top: 10px;
                font-family: 'Segoe UI';
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            """
        )
        # 创建布局
        layout = QVBoxLayout(self)

        # 添加提示标签
        self.label = QLabel(prompt)
        self.label.setFont(QFont("Segoe UI", 14))  # 设置标签字体和大小
        layout.addWidget(self.label)

        # 添加整数输入框
        self.spin_box = QSpinBox()
        self.spin_box.setRange(min_value, max_value)
        self.spin_box.setFont(QFont("Segoe UI", 14))  # 设置输入框字体和大小
        layout.addWidget(self.spin_box)

        # 添加确认按钮
        self.confirm_button = QPushButton("确定")
        self.confirm_button.setFont(QFont("Segoe UI", 14))  # 设置按钮字体和大小
        self.confirm_button.clicked.connect(self.accept)
        layout.addWidget(self.confirm_button)

    def get_value(self):
        """获取用户输入的值"""
        return self.spin_box.value()
    

class CustomTextInputDialog(QDialog):
    def __init__(self, title, prompt, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(300)

        # 设置窗口样式
        self.setStyleSheet(
            """
            QDialog {
                background-color: #f5f5f5;
                border-radius: 8px;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-size: 14px;
            }
            QLabel {
                font-size: 14px;
                color: #333;
                margin-bottom: 10px;
                font-weight: bold;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: white;
                font-size: 14px;
                font-family: 'Segoe UI', 'Arial', sans-serif;
            }
            QPushButton {
                padding: 10px;
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                margin-top: 10px;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            """
        )

        # 创建布局
        layout = QVBoxLayout(self)

        # 添加提示标签
        self.label = QLabel(prompt)
        self.label.setFont(QFont("Segoe UI", 12))  # 设置标签字体和大小
        layout.addWidget(self.label)

        # 添加文本输入框
        self.line_edit = QLineEdit()
        self.line_edit.setFont(QFont("Segoe UI", 12))  # 设置输入框字体和大小
        layout.addWidget(self.line_edit)

        # 添加确认按钮
        self.confirm_button = QPushButton("确定")
        self.confirm_button.setFont(QFont("Segoe UI", 12))  # 设置按钮字体和大小
        self.confirm_button.clicked.connect(self.accept)
        layout.addWidget(self.confirm_button)

    def get_text(self):
        """获取用户输入的文本"""
        return self.line_edit.text()
    
# 将数字转换成金额
def number_to_rmb_upper(amount):
    # 定义数字对应的大写字符
    digit_to_char = {
        '0': '零', '1': '壹', '2': '贰', '3': '叁', '4': '肆',
        '5': '伍', '6': '陆', '7': '柒', '8': '捌', '9': '玖'
    }
    
    # 定义单位
    units = ['', '拾', '佰', '仟']
    higher_units = ['', '万', '亿']
    decimal_units = ['角', '分']
    
    # 分割整数部分和小数部分
    if '.' in amount:
        integer_part, decimal_part = amount.split('.')
    else:
        integer_part, decimal_part = amount, ''
    
    # 处理整数部分
    def convert_integer_part(integer_str):
        length = len(integer_str)
        result = []
        for i, char in enumerate(integer_str):
            if char == '0':
                # 处理连续的零
                if result and result[-1] != '零':
                    result.append('零')
            else:
                result.append(digit_to_char[char])
                result.append(units[(length - i - 1) % 4])
            # 添加高单位（万、亿）
            if (length - i - 1) % 4 == 0 and (length - i - 1) != 0:
                result.append(higher_units[(length - i - 1) // 4])
        # 去除末尾的零
        while result and result[-1] == '零':
            result.pop()
        return ''.join(result)
    
    # 处理小数部分
    def convert_decimal_part(decimal_str):
        result = []
        for i, char in enumerate(decimal_str):
            if i >= len(decimal_units):
                break
            if char != '0':
                result.append(digit_to_char[char])
                result.append(decimal_units[i])
        return ''.join(result)
    
    # 转换整数部分
    integer_upper = convert_integer_part(integer_part)
    if not integer_upper:
        integer_upper = '零'
    
    # 转换小数部分
    decimal_upper = convert_decimal_part(decimal_part)
    
    # 组合结果
    if decimal_upper:
        return f"{integer_upper}元{decimal_upper}"
    else:
        return f"{integer_upper}元整"


# 用逗号分割数字
def format_number_with_commas(num):
    # 输入的字符长度大于10，则不可能是数字
    if len(num) >= 8:
        return num
    
    try:
        num = float(num)
    except Exception as e:
        return num
    return "{:,}".format(num)