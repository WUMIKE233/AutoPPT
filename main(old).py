#近期优化：背景裁切问题，下拉表问题
from PyQt5 import QtCore
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                            QRadioButton, QFileDialog, QComboBox, QProgressBar, QFrame, QLineEdit, QMessageBox,
                            QDialog, QTextBrowser,)  # 添加 QDialog 和 QTextBrowser
from PyQt5.QtGui import QPixmap, QFont, QIcon, QPainterPath, QRegion
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QSize, QRectF
import sys
import os
from datetime import datetime
import seaborn as sns
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import logging
from PPT import PPTApp
import random
import matplotlib as mpl
from matplotlib.font_manager import FontProperties

# 配置中文字体
try:
    font_path = 'C:/Windows/Fonts/simhei.ttf'  # Windows系统字体路径
    if os.path.exists(font_path):
        font_prop = FontProperties(fname=font_path)
        mpl.rcParams['font.family'] = font_prop.get_name()
    else:
        mpl.rcParams['font.family'] = 'SimHei'
except Exception as e:
    logging.warning(f"字体设置失败: {str(e)}")
mpl.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
# 日志配置
logging.basicConfig(
    filename=f'chart_gen_{datetime.now().strftime("%Y%m%d")}.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

class ChartWorker(QThread):
    status_update = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, file_path, chart_type, output_path):
        super().__init__()
        self.file_path = file_path
        self.chart_type = chart_type
        self.output_path = output_path
        self.cancelled = False

    def run(self):
        try:
            self.status_update.emit("正在读取 Excel 文件...")
            df = pd.read_excel(self.file_path)
            if df.empty:
                raise ValueError("Excel 文件为空")

            self.status_update.emit("正在生成图表...")
            plt.figure(figsize=(10, 6))
            # 新增图表类型判断
            if self.chart_type == "饼状图":
                df.iloc[:, 1].value_counts().plot(kind='pie', autopct='%1.1f%%')
                plt.title(f"{df.columns[1]} 分布")

            elif self.chart_type == "柱状图":
                df.plot(kind='bar', x=df.columns[0], y=df.columns[1])
                plt.title(f"{df.columns[1]} 柱状图")

            elif self.chart_type == "折线图":
                df.plot(kind='line', x=df.columns[0], y=df.columns[1])
                plt.title(f"{df.columns[1]} 趋势")

            # 新增散点图
            elif self.chart_type == "散点图":
                if len(df.columns) < 2:
                    raise ValueError("生成散点图需要至少两列数据")
                df.plot(kind='scatter', x=df.columns[0], y=df.columns[1])
                plt.title(f"{df.columns[0]} vs {df.columns[1]} 散点图")

            # 新增热力图
            elif self.chart_type == "热力图":
                numeric_df = df.select_dtypes(include=['number'])
                if numeric_df.empty:
                    raise ValueError("没有数值型数据用于生成热力图")
                corr_matrix = numeric_df.corr()
                sns.heatmap(corr_matrix, annot=True, cmap='coolwarm')
                plt.title("热力图 - 相关性矩阵")

            # 新增直方图
            elif self.chart_type == "直方图":
                if len(df.columns) < 1:
                    raise ValueError("没有可用于生成直方图的数据")
                df[df.columns[1]].plot(kind='hist', bins=10, edgecolor='black')
                plt.title(f"{df.columns[1]} 分布直方图")
                plt.xlabel(df.columns[1])
                plt.ylabel("频数")

            # 新增气泡图
            elif self.chart_type == "气泡图":
                if len(df.columns) < 3:
                    raise ValueError("生成气泡图需要至少三列数据")
                plt.scatter(df.iloc[:, 0], df.iloc[:, 1], s=df.iloc[:, 2] * 50, alpha=0.6)
                plt.title(f"气泡图：{df.columns[0]} vs {df.columns[1]}")
                plt.xlabel(df.columns[0])
                plt.ylabel(df.columns[1])
                plt.grid(True)

            plt.tight_layout()

            if self.cancelled:
                self.finished.emit("图表生成已取消")
                return

            self.status_update.emit("正在保存图表...")
            file_ext = os.path.splitext(self.output_path)[1][1:].lower()
            plt.savefig(self.output_path, dpi=300, format=file_ext)
            plt.close()
            self.finished.emit(f"图表已保存至:\n{self.output_path}")
        except Exception as e:
            self.error.emit(f"生成图表失败: {str(e)}")

    def cancel(self):
        self.cancelled = True

class ChartApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.running = False
        self.selected_folder = r"D:\output Charts"  # 修正变量名和路径
        self.worker = None
        self.logo_path = "./data_init/logo.png"
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.format_combo = QComboBox()  # 提前定义
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("智能图表生成系统 v1.0")
        self.setGeometry(100, 100, 1200, 900)

        # 中央部件设置渐变背景和圆角
        central_widget = QWidget(self)
        central_widget.setStyleSheet("""
            background-color: qradialgradient(cx: 0%, cy: 100%, radius: 100%, stop: 0 #00CED1, stop: 1 #87CEEB);
            border-radius: 50px;
            border: 1px solid #AAAAAA;
        """)
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 标题栏
        title_bar = QWidget(central_widget)
        title_bar.setFixedHeight(140)
        title_bar.setStyleSheet("background: transparent; border: none;")
        title_bar_layout = QHBoxLayout(title_bar)
        title_bar_layout.setContentsMargins(0, 0, 0, 0)

        # Logo
        self.logo_label = QLabel(title_bar)
        self.logo_label.setFixedSize(197, 139)
        self.update_logo()

        # 标题
        self.title_label = QLabel("智能图表生成系统 v1.0", title_bar)
        self.title_label.setFont(QFont("微软雅黑", 24, QFont.Bold))
        self.title_label.setStyleSheet("color: #333333; background: transparent;")
        self.title_label.setAlignment(Qt.AlignCenter)

        # 关闭按钮
        close_btn = QPushButton("×", title_bar)
        close_btn.setFont(QFont("微软雅黑", 14, QFont.Bold))
        close_btn.setFixedSize(30, 30)
        close_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #666666;
                border: none;
            }
            QPushButton:hover {
                color: #FF0000;
            }
        """)
        close_btn.clicked.connect(self.close)

        title_bar_layout.addWidget(self.logo_label)
        title_bar_layout.addWidget(self.title_label)
        title_bar_layout.addStretch()
        title_bar_layout.addWidget(close_btn)
        main_layout.addWidget(title_bar)

        # 文件上传区
        file_frame = QFrame(central_widget)
        file_frame.setStyleSheet("background: transparent;")
        file_layout = QHBoxLayout(file_frame)
        file_label = QLabel("选择 Excel 文件:", file_frame)
        file_label.setFont(QFont("微软雅黑", 12))
        self.file_path = QLineEdit(file_frame)
        self.file_path.setReadOnly(True)
        self.file_path.setFixedHeight(60)
        file_btn = QPushButton("选择文件", file_frame)
        file_btn.setFixedHeight(60)
        file_btn.setFont(QFont("微软雅黑", 12))
        file_btn.clicked.connect(self.load_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(file_btn)
        main_layout.addWidget(file_frame)

        # 图表类型选择
        chart_frame = QFrame(central_widget)
        chart_frame.setStyleSheet("background: transparent;")
        chart_layout = QHBoxLayout(chart_frame)
        chart_label = QLabel("图表类型:", chart_frame)
        chart_label.setFont(QFont("微软雅黑", 12))
        self.chart_combo = QComboBox(chart_frame)
        self.chart_combo.addItems([
            "饼状图", "柱状图", "折线图",
            "散点图", "热力图", "直方图", "气泡图"
        ])
        self.chart_combo.setFixedHeight(60)
        self.chart_combo.setFont(QFont("微软雅黑", 12))
        # 设置下拉表的样式
        self.chart_combo.setStyleSheet("""
            QComboBox {
                background-color: rgba(255, 255, 255, 0.9);
                color: #333333;
                border: 1px solid #AAAAAA;
                border-radius: 10px;
                padding: 5px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: url(./data_init/down_arrow.png);  # 可选：添加下拉箭头图标
                width: 10px;
                height: 10px;
            }
            QComboBox QAbstractItemView {
                background-color: rgba(255, 255, 255, 0.95);
                color: #333333;
                selection-background-color: rgba(53, 142, 255, 0.8);
                selection-color: white;
                border: 1px solid #AAAAAA;
                border-radius: 5px;
            }
        """)
        chart_layout.addWidget(chart_label)
        chart_layout.addWidget(self.chart_combo)
        main_layout.addWidget(chart_frame)

        # 输出设置
        output_frame = QFrame(central_widget)
        output_frame.setStyleSheet("background: transparent;")
        output_layout = QHBoxLayout(output_frame)
        folder_label = QLabel("输出文件夹:", output_frame)
        folder_label.setFont(QFont("微软雅黑", 12))
        self.folder_input = QLineEdit(self.selected_folder, output_frame)
        self.folder_input.setFixedHeight(60)
        folder_btn = QPushButton("选择文件夹", output_frame)
        folder_btn.setFixedHeight(60)
        folder_btn.setFont(QFont("微软雅黑", 12))
        folder_btn.clicked.connect(self.select_folder)
        output_label = QLabel("输出文件:", output_frame)
        output_label.setFont(QFont("微软雅黑", 12))
        self.output_path = QLineEdit(self.get_default_output_path(), output_frame)
        self.output_path.setReadOnly(True)
        self.output_path.setFixedHeight(60)
        output_layout.addWidget(folder_label)
        output_layout.addWidget(self.folder_input)
        output_layout.addWidget(folder_btn)
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_path)
        main_layout.addWidget(output_frame)

        # 输出格式选择
        self.format_combo.currentTextChanged.connect(self.update_output_path_extension)

        format_frame = QFrame(central_widget)
        format_frame.setStyleSheet("background: transparent;")
        format_layout = QHBoxLayout(format_frame)
        format_label = QLabel("输出格式:", format_frame)
        format_label.setFont(QFont("微软雅黑", 12))
        self.format_combo = QComboBox(format_frame)
        self.format_combo.addItems(["PNG", "PDF", "SVG"])
        self.format_combo.setFixedHeight(60)
        self.format_combo.setFont(QFont("微软雅黑", 12))
        # 设置输出格式下拉表的样式
        self.format_combo.setStyleSheet("""
            QComboBox {
                background-color: rgba(255, 255, 255, 0.9);
                color: #333333;
                border: 1px solid #AAAAAA;
                border-radius: 10px;
                padding: 5px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: url(./data_init/down_arrow.png);  # 可选：添加下拉箭头图标
                width: 10px;
                height: 10px;
            }
            QComboBox QAbstractItemView {
                background-color: rgba(255, 255, 255, 0.95);
                color: #333333;
                selection-background-color: rgba(53, 142, 255, 0.8);
                selection-color: white;
                border: 1px solid #AAAAAA;
                border-radius: 5px;
            }
        """)
        format_layout.addWidget(format_label)
        format_layout.addWidget(self.format_combo)
        main_layout.addWidget(format_frame)

        # 控制按钮区
        control_frame = QFrame(central_widget)
        control_frame.setStyleSheet("background: transparent;")
        control_layout = QHBoxLayout(control_frame)
        self.progress = QProgressBar(control_frame)
        self.progress.setRange(0, 0)
        self.progress.hide()
        self.generate_btn = QPushButton("生成图表", control_frame)
        self.generate_btn.setFixedHeight(60)
        self.generate_btn.setFont(QFont("微软雅黑", 12, QFont.Bold))
        self.generate_btn.clicked.connect(self.start_generate)
        self.cancel_btn = QPushButton("取消", control_frame)
        self.cancel_btn.setFixedHeight(60)
        self.cancel_btn.setFont(QFont("微软雅黑", 12, QFont.Bold))
        self.cancel_btn.clicked.connect(self.cancel_generate)
        self.cancel_btn.setEnabled(False)
        self.status_label = QLabel("就绪", control_frame)
        self.status_label.setStyleSheet("color: white; background: transparent;")
        self.status_label.setFont(QFont("微软雅黑", 12))
        control_layout.addWidget(self.progress)
        control_layout.addWidget(self.generate_btn)
        control_layout.addWidget(self.cancel_btn)
        control_layout.addWidget(self.status_label)
        main_layout.addWidget(control_frame)

        # 统一按钮样式
        for btn in [file_btn, folder_btn, self.generate_btn, self.cancel_btn]:
            btn.setStyleSheet("""
                QPushButton {
                    background: rgba(53, 142, 255, 0.9);
                    color: white;
                    border: none;
                    border-radius: 10px;
                    padding: 10px 20px;
                }
                QPushButton:disabled {background: rgba(150, 150, 150, 0.6);}
                QPushButton:hover {background: rgba(53, 142, 255, 1);}
            """)

        # 添加动态标题颜色效果
        self.update_title_color()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        path = QPainterPath()
        path.addRoundedRect(QRectF(0, 0, self.width(), self.height()), 50, 50)
        region = QRegion(path.toFillPolygon().toPolygon())
        self.setMask(region)

    def update_logo(self):
        if os.path.exists(self.logo_path):
            pixmap = QPixmap(self.logo_path)
            pixmap = pixmap.scaled(self.logo_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(pixmap)
            self.logo_label.setStyleSheet("background: transparent;")
        else:
            logging.warning(f"Logo文件不存在: {self.logo_path}")

    def update_title_color(self):
        r, g, b = [random.randint(0, 255) for _ in range(3)]
        self.title_label.setStyleSheet(f"color: rgb({r}, {g}, {b}); background: transparent;")
        QTimer.singleShot(500, self.update_title_color)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_pos = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() & QtCore.Qt.LeftButton and self.drag_pos:
            self.move(event.globalPos() - self.drag_pos)
            event.accept()

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if path:
            self.file_path.setText(path)
            self.output_path.setText(self.get_default_output_path())

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", self.selected_folder)
        if folder:
            self.selected_folder = folder
            self.folder_input.setText(folder)
            self.output_path.setText(self.get_default_output_path())

    def get_default_output_path(self):
        folder = os.path.abspath(self.folder_input.text())
        selected_format = self.format_combo.currentText().lower()
        filename = f"chart_{datetime.now().strftime('%Y%m%d%H%M')}.{selected_format}"
        os.makedirs(folder, exist_ok=True)
        return os.path.join(folder, filename)

    def update_output_path_extension(self, text):
        if self.output_path.text():
            current_path = self.output_path.text()
            base_name = os.path.split(current_path)[1].rsplit('.', 1)[0]  # 修正扩展名处理
            new_extension = text.lower()
            new_path = f"{base_name}.{new_extension}"
            self.output_path.setText(new_path)  # 修正方法名

    def start_generate(self):
        if not self.running and self.file_path.text():
            self.running = True
            self.progress.show()
            self.status_label.setText("生成图表中...")
            self.generate_btn.setEnabled(False)
            self.cancel_btn.setEnabled(True)
            self.worker = ChartWorker(self.file_path.text(), self.chart_combo.currentText(), self.output_path.text())
            self.worker.status_update.connect(self.update_status)
            self.worker.finished.connect(self.on_generate_finished)
            self.worker.error.connect(self.on_generate_error)
            self.worker.start()

    def cancel_generate(self):
        if self.running and self.worker:
            self.worker.cancel()
            self.status_label.setText("取消中...")

    def update_status(self, status):
        self.status_label.setText(status)

    def on_generate_finished(self, message):
        self.running = False
        self.progress.hide()
        self.status_label.setText("就绪")
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        QMessageBox.information(self, "成功", message)

    def on_generate_error(self, message):
        self.running = False
        self.progress.hide()
        self.status_label.setText("就绪")
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        QMessageBox.critical(self, "错误", message)

# HelpWindow 类定义
class HelpWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        self.setGeometry(100, 100, 800, 600)
        self.setWindowTitle("Help - Intelligent Generation Tool")
        self.text_browser = QTextBrowser(self)
        self.text_browser.setHtml(self.get_help_text())
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.text_browser)
        self.setStyleSheet("background-color: white;")
        self.setLayout(main_layout)

    def get_help_text(self):
        return """
        <h1>欢迎使用智能生成工具</h1>
        <p>该应用提供两个主要功能：</p>
        <ul>
            <li>从文本输入或 Markdown 文件生成 PowerPoint 演示文稿（PPT）。</li>
            <li>从 Excel 文件生成图表。</li>
        </ul>
        <h2>生成 PPT</h2>
        <p>要生成一个 PPT：</p>
        <ol>
            <li>选择“生成 PPT”按钮。</li>
            <li>选择文本输入或文件上传来提供您的内容。</li>
            <li>输入您的主题或加载一个 Markdown 文件。</li>
            <li>从可用模板中选择一个模板。</li>
            <li>选择输出文件夹和文件名。</li>
            <li>生成大纲，然后确认生成 PPT。</li>
        </ol>
        <h2>生成图表</h2>
        <p>要生成一个图表：</p>
        <ol>
            <li>选择“生成图表”按钮。</li>
            <li>加载一个 Excel 文件。</li>
            <li>选择您想生成的图表类型。</li>
            <li>选择输出文件夹和文件名。</li>
            <li>点击“生成图表”来创建图表。</li>
        </ol>
        <h2>一般使用</h2>
        <p>- 该应用具有用户友好的界面，带有清晰的标签和按钮。</p>
        <p>- 确保您有文件操作的必要权限。</p>
        <p>- 遇到任何问题，请参考日志文件或联系支持。</p>
        """

# MainApp 类定义
class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ppt_app = None
        self.chart_app = None
        self.logo_path = "./data_init/logo.png"
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.init_ui()
        self.is_full_screen = False

    def init_ui(self):
        self.setWindowTitle("智能生成工具")
        self.setGeometry(100, 100, 1200, 900)

        # 中央部件设置渐变背景和圆角
        central_widget = QWidget(self)
        central_widget.setStyleSheet("""
            background-color: qradialgradient(cx: 0%, cy: 100%, radius: 100%, stop: 0 #00CED1, stop: 1 #87CEEB);
            border-radius: 50px;
            border: 0px solid #AAAAAA;
        """)
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 标题栏
        title_bar = QWidget(central_widget)
        title_bar.setFixedHeight(140)
        title_bar.setStyleSheet("background: transparent; border: none;")
        title_bar_layout = QHBoxLayout(title_bar)
        title_bar_layout.setContentsMargins(0, 0, 0, 0)

        self.logo_label = QLabel(title_bar)
        self.logo_label.setFixedSize(197, 139)
        self.update_logo()

        self.title_label = QLabel("智能生成工具", title_bar)
        self.title_label.setFont(QFont("微软雅黑", 24, QFont.Bold))
        self.title_label.setStyleSheet("color: #333333; background: transparent;")
        self.title_label.setAlignment(Qt.AlignCenter)

        close_btn = QPushButton("×", title_bar)
        close_btn.setFont(QFont("微软雅黑", 14, QFont.Bold))
        close_btn.setFixedSize(30, 30)
        close_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #666666;
                border: none;
            }
            QPushButton:hover {
                color: #FF0000;
            }
        """)
        close_btn.clicked.connect(self.close)

        # 添加“Help”按钮
        self.help_button = QPushButton("Help", title_bar)
        self.help_button.setFixedSize(80, 30)
        self.help_button.setFont(QFont("微软雅黑", 12))
        self.help_button.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #666666;
                border: none;
            }
            QPushButton:hover {
                color: #0000FF;
            }
        """)
        self.help_button.clicked.connect(self.show_help_window)

        title_bar_layout.addWidget(self.logo_label)
        title_bar_layout.addWidget(self.title_label)
        title_bar_layout.addStretch()
        title_bar_layout.addWidget(self.help_button)  # 添加 Help 按钮
        title_bar_layout.addWidget(close_btn)
        main_layout.addWidget(title_bar)

        # 添加功能描述文本框，设置框架背景为透明
        desc_frame = QFrame(central_widget)
        desc_frame.setStyleSheet("background: transparent;")
        desc_layout = QHBoxLayout(desc_frame)
        desc_label = QLabel("本系统支持根据描述生成PPT，或将Excel表格导入后进行可视化生成图表", desc_frame)
        desc_label.setFont(QFont("微软雅黑", 14))
        desc_label.setStyleSheet("""
            color: #333333;
            background: transparent;
            border: 1px solid #AAAAAA;
            border-radius: 5px;
            padding: 10px;
        """)
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setWordWrap(True)
        desc_layout.addWidget(desc_label)
        main_layout.addWidget(desc_frame)

        # 按钮区，设置框架背景为透明
        btn_frame = QFrame(central_widget)
        btn_frame.setStyleSheet("background: transparent;")
        btn_layout = QHBoxLayout(btn_frame)
        btn_layout.setSpacing(20)

        # 创建 PPT 按钮和标签组合
        ppt_container = QWidget()
        ppt_container.setStyleSheet("background: transparent; border: none;")
        ppt_layout = QVBoxLayout(ppt_container)
        ppt_layout.setAlignment(Qt.AlignCenter)
        ppt_layout.setSpacing(10)

        self.ppt_btn = QPushButton(self)
        self.ppt_btn.setFixedSize(590, 390)
        ppt_bg_path = os.path.join(os.path.dirname(__file__), "data_init", "ppt_button_bg.jpg")
        if not os.path.exists(ppt_bg_path):
            logging.error(f"PPT按钮背景图片不存在: {ppt_bg_path}")
        else:
            logging.info(f"PPT按钮背景图片路径: {ppt_bg_path}")
            ppt_pixmap = QPixmap(ppt_bg_path)
            if not ppt_pixmap.isNull():
                self.ppt_btn.setIcon(QIcon(ppt_pixmap))
                self.ppt_btn.setIconSize(QSize(590, 390))
                logging.info("PPT按钮使用 QPixmap 设置图标成功")
            else:
                logging.error(f"PPT按钮图片加载失败: {ppt_bg_path}")
        ppt_bg_path_fixed = ppt_bg_path.replace('\\', '/')
        ppt_bg_url = f"file:///{ppt_bg_path_fixed}"
        logging.info(f"PPT按钮背景URL: {ppt_bg_url}")
        self.ppt_btn.setStyleSheet(
            "QPushButton {"
            f"background-image: url({ppt_bg_url});"
            "background-position: center;"
            "background-repeat: no-repeat;"
            "background-size: contain;"
            "border: none;"
            "border-radius: 10px;"
            "}"
            "QPushButton:hover {"
            f"background-image: url({ppt_bg_url});"
            "background-color: rgba(50, 50, 50, 0.5);"
            "background-position: center;"
            "background-repeat: no-repeat;"
            "background-size: contain;"
            "}"
        )
        self.ppt_btn.clicked.connect(self.open_ppt_app)

        ppt_label = QLabel("生成 PPT", self)
        ppt_label.setFont(QFont("微软雅黑", 16, QFont.Bold))
        ppt_label.setStyleSheet("color: #000000; background: transparent; border: none;")
        ppt_label.setAlignment(Qt.AlignCenter)

        ppt_layout.addWidget(self.ppt_btn)
        ppt_layout.addWidget(ppt_label)

        # 创建 Chart 按钮和标签组合
        chart_container = QWidget()
        chart_container.setStyleSheet("background: transparent; border: none;")
        chart_layout = QVBoxLayout(chart_container)
        chart_layout.setAlignment(Qt.AlignCenter)
        chart_layout.setSpacing(10)

        self.chart_btn = QPushButton(self)
        self.chart_btn.setFixedSize(590, 390)
        chart_bg_path = os.path.join(os.path.dirname(__file__), "data_init", "chart_button_bg.jpg")
        if not os.path.exists(chart_bg_path):
            logging.error(f"Chart按钮背景图片不存在: {chart_bg_path}")
        else:
            logging.info(f"Chart按钮背景图片路径: {chart_bg_path}")
            chart_pixmap = QPixmap(chart_bg_path)
            if not chart_pixmap.isNull():
                self.chart_btn.setIcon(QIcon(chart_pixmap))
                self.chart_btn.setIconSize(QSize(590, 390))
                logging.info("Chart按钮使用 QPixmap 设置图标成功")
            else:
                logging.error(f"Chart按钮图片加载失败: {chart_bg_path}")
        chart_bg_path_fixed = chart_bg_path.replace('\\', '/')
        chart_bg_url = f"file:///{chart_bg_path_fixed}"
        logging.info(f"Chart按钮背景URL: {chart_bg_url}")
        self.chart_btn.setStyleSheet(
            "QPushButton {"
            f"background-image: url({chart_bg_url});"
            "background-position: center;"
            "background-repeat: no-repeat;"
            "background-size: contain;"
            "border: none;"
            "border-radius: 10px;"
            "}"
            "QPushButton:hover {"
            f"background-image: url({chart_bg_url});"
            "background-color: rgba(50, 50, 50, 0.5);"
            "background-position: center;"
            "background-repeat: no-repeat;"
            "background-size: contain;"
            "}"
        )
        self.chart_btn.clicked.connect(self.open_chart_app)

        chart_label = QLabel("生成图表", self)
        chart_label.setFont(QFont("微软雅黑", 16, QFont.Bold))
        chart_label.setStyleSheet("color: #000000; background: transparent; border: none;")
        chart_label.setAlignment(Qt.AlignCenter)

        chart_layout.addWidget(self.chart_btn)
        chart_layout.addWidget(chart_label)

        btn_layout.addWidget(ppt_container)
        btn_layout.addWidget(chart_container)
        main_layout.addWidget(btn_frame)

        # 添加时间控件到右下角，设置框架背景为透明
        time_frame = QFrame(central_widget)
        time_frame.setStyleSheet("background: transparent;")
        time_layout = QHBoxLayout(time_frame)
        time_layout.setAlignment(Qt.AlignRight)
        self.time_label = QLabel(self.get_current_time(), time_frame)
        self.time_label.setFont(QFont("微软雅黑", 12))
        self.time_label.setStyleSheet("color: #333333; background: transparent;")
        time_layout.addWidget(self.time_label)
        main_layout.addStretch()
        main_layout.addWidget(time_frame)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

        self.update_title_color()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 设置主窗口遮罩为圆角矩形
        path = QPainterPath()
        path.addRoundedRect(QRectF(0, 0, self.width(), self.height()), 50, 50)
        region = QRegion(path.toFillPolygon().toPolygon())
        self.setMask(region)

    def get_current_time(self):
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def update_time(self):
        self.time_label.setText(self.get_current_time())

    def update_logo(self):
        if os.path.exists(self.logo_path):
            pixmap = QPixmap(self.logo_path)
            pixmap = pixmap.scaled(self.logo_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(pixmap)
            self.logo_label.setStyleSheet("background: transparent;")
        else:
            logging.warning(f"Logo文件不存在: {self.logo_path}")

    def update_title_color(self):
        r, g, b = [random.randint(0, 255) for _ in range(3)]
        self.title_label.setStyleSheet(f"color: rgb({r}, {g}, {b}); background: transparent;")
        QTimer.singleShot(500, self.update_title_color)

    def open_ppt_app(self):
        print("Opening PPTApp")
        if not self.ppt_app:
            self.ppt_app = PPTApp()
        self.ppt_app.show()

    def open_chart_app(self):
        print("Opening ChartApp")
        if not self.chart_app:
            self.chart_app = ChartApp()
        self.chart_app.show()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_pos = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and self.drag_pos:
            self.move(event.globalPos() - self.drag_pos)
            event.accept()

    def show_help_window(self):
        """显示帮助窗口"""
        help_window = HelpWindow(self)
        help_window.exec_()  # 使用模态窗口显示

# 主程序运行部分（保持不变）
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('./data_init/app_icon.ico'))
    window = MainApp()
    window.show()
    sys.exit(app.exec_())



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('./data_init/app_icon.ico'))
    window = MainApp()
    window.show()
    sys.exit(app.exec_())