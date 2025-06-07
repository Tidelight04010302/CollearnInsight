import sys
import os 
import copy
import numpy as np  
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QSplitter,
                             QPushButton, QLineEdit, QSlider, QVBoxLayout, QHBoxLayout,
                             QFileDialog, QLabel, QComboBox, QListWidget, QTableWidget,
                             QTableWidgetItem, QScrollArea, QHeaderView, QStyleFactory,
                             QCheckBox, QFrame,QSizePolicy,QGridLayout,QToolButton, 
                             QMessageBox,QTabBar,QSpinBox,QInputDialog,QGroupBox)
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtMultimediaWidgets import QVideoWidget
from PyQt5.QtCore import Qt, QUrl, QTimer, pyqtSignal, QMimeData,QByteArray,QSize
from PyQt5.QtGui import QColor, QPalette, QFont,QIcon,QPixmap,QIntValidator
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt  
from matplotlib import colormaps as colormaps
from sklearn.cluster import AgglomerativeClustering 
from sklearn.preprocessing import StandardScaler 
from scipy.cluster.hierarchy import dendrogram, linkage  
from collections import defaultdict
class CombinedApplication(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("协作学习编码与分析工具（CollearnInsight）")
        self.setGeometry(100, 100, 1900, 800)
        # 创建主布局容器
        main_container =QSplitter(Qt.Horizontal)
        tab_container=QSplitter(Qt.Horizontal)
        tab_module_container=QSplitter(Qt.Horizontal)
         # 创建标签页控件
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border:0px solid #FFFFFF;
                background: #f8f8ff;
            }
            QTabBar::tab {
                background: #F8F8FF;
                padding: 8px 12px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 120px;     /* 设置最小宽度 */
                min-height: 30px;     /* 设置最小高度 */
                font: 10pt "黑体";
            }
            QTabBar::tab:selected {
                background: #a6baff;
                border-bottom: 2px solid white;
            }
        """)
        self.tabmodule_widget = QTabWidget()
        self.tabmodule_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #e8e8e8;
                background: #ffffff;
            }
            QTabBar::tab {
                background: #d0d9ff;
                padding: 8px 12px;
                margin-right: 2px;
                border: 1px solid #e8e8e8;                    
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 120px;     /* 设置最小宽度 */
                min-height: 30px;     /* 设置最小高度 */
                font: 10pt "黑体";
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                border-bottom: 2px solid white;
            }
        """)
        # 创建两个功能模块
        self.video_player = VideoPlayerModule()
        self.data_encoder = DataEncoderModule()
        self.cluster_analysis = ClusterAnalysisModule(tabmodule=self.tabmodule_widget)
        self.sequence_analysis = SequenceAnalysisModule()
        # 正确传入 tabmodule 参数
        self.summary_tab = SummaryTab(tabmodule=self.tabmodule_widget)
        # 将功能模块添加到标签页
        # 设置最小宽度替代固定宽度
        
        # 添加“+”按钮用于新增标签页
        self.add_tab_button = QPushButton('+')
        self.add_tab_button.setFixedSize(30, 30)
        self.add_tab_button.setStyleSheet("""
        QPushButton {
        background-color: #d0d9ff;
        font-size: 16px;
        font-weight: bold;
        border-radius: 15px;
            }
        """)
        self.add_tab_button.clicked.connect(self.add_new_tab)
        # 设置为标签页的右上角控件
        # 在CombinedApplication.__init__中调整标签页结构
        self.tabmodule_widget.addTab(self.data_encoder,'编码个体1') 
        self.tabmodule_widget.addTab(self.summary_tab, "数据汇总")
        # 添加标签关闭功能
        self.tabmodule_widget.setTabsClosable(True)
        self.tabmodule_widget.tabCloseRequested.connect(
            self.tabmodule_widget.removeTab
        )
        self.tabmodule_widget.setMovable(False)
        tab_module_container = QWidget()
        module_layout = QVBoxLayout()
        module_layout.addWidget(self.tabmodule_widget)
        module_layout.addWidget(self.add_tab_button)  # 按钮移到内层
        module_layout.setContentsMargins(0, 0, 0, 0)
        tab_module_container.setAutoFillBackground(True)
        p = tab_module_container.palette()
        p.setColor(tab_module_container.backgroundRole(), QColor(248, 248, 255))
        tab_module_container.setPalette(p)
        tab_module_container.setLayout(module_layout)
        # 构建外层标签页内容
        tab_container = QWidget()
        splitter = QSplitter()

        # 将两个控件添加到分割器中
        splitter.addWidget(self.video_player)
        splitter.addWidget(tab_module_container)

        # 设置分割器的拉伸因子（可选）
        self.video_player.setMinimumWidth(650)
        self.video_player.setMaximumWidth(1200)
        tab_module_container.setMinimumWidth(700)
        tab_module_container.setMaximumWidth(1250)
        main_layout = QVBoxLayout()  # 假设这是你的主布局
        main_layout.addWidget(splitter)
        
        main_layout.setSpacing(0)
        tab_container.setAutoFillBackground(True)
        p = tab_container.palette()
        p.setColor(tab_container.backgroundRole(), QColor(248, 248, 255))
        tab_container.setPalette(p)
        tab_container.setLayout(main_layout)
        sequence_analysis_container = QWidget()
        sequence_analysis_layout = QHBoxLayout()
        sequence_analysis_layout.addWidget(self.sequence_analysis)
        sequence_analysis_layout.setSpacing(0)
        sequence_analysis_container.setAutoFillBackground(True)
        p = sequence_analysis_container.palette()
        p.setColor(sequence_analysis_container.backgroundRole(), QColor(248, 248, 255))
        sequence_analysis_container.setPalette(p)
        sequence_analysis_container.setLayout(sequence_analysis_layout)
        self.tab_widget.addTab(tab_container, "数据编码")
        self.tab_widget.addTab(sequence_analysis_container, "序列分析")
        self.tab_widget.addTab(self.cluster_analysis, "聚类分析")
        self.tab_widget.currentChanged.connect(self.on_outer_tab_change)
        main_container.addWidget(self.tab_widget)
        # 添加图标区分层级
        self.tab_widget.setTabIcon(0, QIcon("icons/data.png"))
        self.tabmodule_widget.setTabIcon(0, QIcon("icons/group.png"))
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(248, 248, 255))
        self.setPalette(palette)
        # 设置主窗口中心部件
        self.setCentralWidget(main_container)
    # 修改add_new_tab方法（建议改为内层标签页专用）
    def add_new_tab(self):
        used_indices = []
        for i in range(self.tabmodule_widget.count()):
            tab_text = self.tabmodule_widget.tabText(i)
            if tab_text == "数据汇总":
                continue
            if tab_text.startswith("编码个体"):
                try:
                    index = int(tab_text[4:])  # 提取数字部分
                    used_indices.append(index)
                except:
                    pass
        # 找到最小未使用的序号
        used_indices.sort()
        next_index = 1
        for idx in used_indices:
            if idx == next_index:
                next_index += 1
            elif idx > next_index:
                break
        # 创建新标签页并设置名称
        new_tab = DataEncoderModule()
        self.tabmodule_widget.addTab(new_tab, f"编码个体{next_index}")
        self.move_summary_to_end()
    def on_outer_tab_change(self, index):
        """外层标签页切换时的处理"""
        if self.tab_widget.tabText(index) == "数据编码":
            # 可在此处刷新内层标签页数据
            pass
    def move_summary_to_end(self):
        summary_index = self.tabmodule_widget.indexOf(self.summary_tab)
        if summary_index != -1:
            self.tabmodule_widget.removeTab(summary_index)
            self.tabmodule_widget.addTab(self.summary_tab, "数据汇总")
class SummaryTab(QWidget):
    def __init__(self, parent=None, tabmodule=None):
        super().__init__(parent)
        if tabmodule is None:
            raise ValueError("必须传入 tabmodule 参数")
        self.tabmodule = tabmodule  # 保存对内层标签页的引用
        # 初始化UI
        self.initUI()
        # 初始化时加载数据
        self.generate_summary_data()
        # 监听标签页变化
        self.tabmodule.currentChanged.connect(self.generate_summary_data)
    def initUI(self):
        """初始化汇总页面UI"""
        layout = QVBoxLayout()
        # 创建表格
        self.summary_table = QTableWidget()
        self.summary_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #E0E0E0;
                selection-background-color: #d0d9ff;
            }
        """)
        time_notice_label = QLabel("注意：编码个体的时间需要保持一致")
        time_notice_label.setFont(QFont('黑体', 12))
        time_notice_label.setStyleSheet("font-size: 15px;")
        layout.addWidget(time_notice_label)
        # 创建输出按钮
        output_btn = QPushButton('输出汇总数据')
        output_btn.setFont(QFont('黑体', 12))
        output_btn.setStyleSheet("font-size: 16px;")
        output_btn.setFixedSize(120, 50)
        output_btn.clicked.connect(self.output_summary_data)
        # 布局组装
        layout.addWidget(self.summary_table)
        layout.addWidget(output_btn, 0, Qt.AlignRight)
        self.setLayout(layout)
    def generate_summary_data(self):
        """生成汇总数据"""
        # 获取所有非汇总标签页
        data_pages = []
        for i in range(self.tabmodule.count()):
            widget = self.tabmodule.widget(i)
            if isinstance(widget, DataEncoderModule):
                data_pages.append(widget)
        if not data_pages:
            self.summary_table.setRowCount(0)
            return
        # 收集数据
        all_data = []
        columns = ["Time"]  # 固定时间列
        max_valid_rows = 0
        best_page = None
        best_time_data = []
        for page in data_pages:
            time_data = self.get_time_column_rows(page)
            valid_count = len(time_data)
            if valid_count > max_valid_rows:
                max_valid_rows = valid_count
                best_page = page
                best_time_data = time_data
        # 填充时间列
        if best_page and best_time_data:
            for row_idx, time_str in enumerate(best_time_data):
                if row_idx >= len(all_data):
                    all_data.append({})
                try:
                    # 支持基本时间格式
                    if ':' in time_str:
                        m, s = map(int, time_str.split(':'))
                        all_data[row_idx]["Time"] = str(m * 60 + s)
                    elif time_str.isdigit():
                        all_data[row_idx]["Time"] = time_str
                    else:
                        all_data[row_idx]["Time"] = str(0)
                except:
                    all_data[row_idx]["Time"] = 0
        # 处理所有数据页面
        for page_idx, page in enumerate(data_pages):
            # 获取表格数据
            table = page.result_table
            row_count = table.rowCount()
            # 处理当前页面数据
            for row_idx in range(row_count):
                # 如果是第一行，初始化该行
                if row_idx >= len(all_data):
                    all_data.append({})
                current_dim_index = 0  # 每个页面独立的维度计数器
                # 处理编码维度列
                for col_idx in range(table.columnCount()):
                    header = table.horizontalHeaderItem(col_idx).text()
                    # 跳过序号、时间和注释列
                    if '序号' in header or '时间' in header or '注释' in header:
                        continue
                    # 构造新列名 S{页面序号}_Eoding{维度序号}
                    current_dim_index += 1  # 每个有效列递增
                    new_header = f"S{page_idx + 1}Eoding_{current_dim_index}"
                    # 添加列头
                    if new_header not in columns:
                        columns.append(new_header)
                    # 获取数据项
                    if table.item(row_idx, col_idx):
                        all_data[row_idx][new_header] = table.item(row_idx, col_idx).text()
        # 更新表格
        self.update_summary_table(columns, all_data)
    def update_summary_table(self, columns, data):
        """更新汇总表格"""
        self.summary_table.setColumnCount(len(columns))
        self.summary_table.setHorizontalHeaderLabels(columns)
        self.summary_table.setRowCount(len(data))
        # 填充数据
        for row_idx, row_data in enumerate(data):
            for col_idx, col_name in enumerate(columns):
                value = row_data.get(col_name, '')
                self.summary_table.setItem(row_idx, col_idx, QTableWidgetItem(value))
        # 自动调整列宽
        self.summary_table.resizeColumnsToContents()
        self.summary_table.resizeRowsToContents()
    def output_summary_data(self):
        """输出汇总数据到Excel"""
        if self.summary_table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "没有可输出的数据！")
            return
        # 获取保存路径
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存汇总数据", "", "Excel 文件 (*.xlsx);;所有文件 (*)"
        )
        if not file_path:
            return
        try:
            # 准备DataFrame
            df = pd.DataFrame()
            # 提取数据
            for col_idx in range(self.summary_table.columnCount()):
                col_name = self.summary_table.horizontalHeaderItem(col_idx).text()
                df[col_name] = [self.summary_table.item(row_idx, col_idx).text() 
                               for row_idx in range(self.summary_table.rowCount())]
            # 导出到Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='汇总', index=False)
                worksheet = writer.sheets['汇总']
                # 自动调整列宽
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    worksheet.column_dimensions[column].width = max_length + 2
                    # 居中对齐
                    for cell in col:
                        cell.alignment = Alignment(horizontal='center')
                # 标题加粗
                for cell in worksheet["1:1"]:
                    cell.font = Font(bold=True)
            QMessageBox.information(self, "成功", "数据已成功导出！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")
    def get_time_column_rows(self, page):
        """统计某个 DataEncoderModule 页面中时间列非空的行数"""
        table = page.result_table
        row_count = table.rowCount()
        valid_rows = []
        for row in range(row_count):
            time_item = table.item(row, 1)  # 时间列在第1列
            if time_item and time_item.text().strip():
                valid_rows.append(time_item.text())
        return valid_rows
class VideoPlayerModule(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.current_position = 0
        self.segment_duration = 10
        self.current_speed = 1.0
        self.current_file_path = ""
    def initUI(self):
        layout = QHBoxLayout(self)
        # 文件列表区域
        file_list_layout = QVBoxLayout()
        self.file_label = QLabel("视频列表")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setFont(QFont('黑体', 12))
        self.file_list = QListWidget()
        self.open_button = QPushButton('打开')
        self.open_button.setFont(QFont('黑体', 10))
        self.open_button.setFixedHeight(30)
        file_list_layout.addWidget(self.file_label)
        file_list_layout.addWidget(self.file_list)
        file_list_layout.addWidget(self.open_button)
        file_list_container = QWidget()
        file_list_container.setLayout(file_list_layout)
        file_list_container.setFixedWidth(100)
        # 主内容区域
        content_layout = QVBoxLayout()
        # 视频显示区域
        self.video_widget = QVideoWidget()
        self.video_widget.setStyleSheet("""
            border: 2px solid white;
            background-color: '#a6baff';
        """)
        self.media_player = QMediaPlayer(None, QMediaPlayer.VideoSurface)
        self.media_player.setVideoOutput(self.video_widget)
        # 控制按钮区域
        self.play_button = QPushButton('播放')
        self.play_button.setFont(QFont('黑体', 10))
        self.play_button.setFixedSize(60, 35)
        self.replay_button = QPushButton('重播')
        self.replay_button.setFont(QFont('黑体', 10))
        self.replay_button.setFixedSize(60, 35)
        self.jump_inputmin = QLineEdit()
        self.jump_inputmin.setFont(QFont('Times New Roman', 11))
        self.jump_inputmin.setPlaceholderText("分")
        self.jump_inputmin.setFixedSize(35, 35)
        self.jump_inputlabel = QLabel(':')
        self.jump_inputlabel.setFont(QFont('Times New Roman', 11))
        self.jump_inputlabel.setStyleSheet("color: black;")
        self.jump_inputlabel.setFixedSize(5, 35)
        self.jump_input = QLineEdit()
        self.jump_input.setFont(QFont('Times New Roman', 11))
        self.jump_input.setPlaceholderText("秒")
        self.jump_input.setFixedSize(35, 35)
        # 添加整数验证器，限制输入范围为 0~59
        self.jump_inputmin.setValidator(QIntValidator(0, 59, self))
        self.jump_input.setValidator(QIntValidator(0, 59, self))
        self.jump_button = QPushButton('跳转')
        self.jump_button.setFont(QFont('黑体', 10))
        self.jump_button.setFixedSize(50, 35)
        self.speed_label = QLabel("播放速度：")
        self.speed_label.setAlignment(Qt.AlignVCenter|Qt.AlignRight)
        self.speed_label.setFont(QFont('宋体', 12))
        self.speed_label.setStyleSheet("color: black;")
        self.speed_combo = QComboBox()
        self.speed_combo.setFont(QFont('Times New Roman', 11))
        self.speed_combo.setFixedSize(75, 35)
        # 进度控制区域
        self.progress_slider = QSlider(Qt.Horizontal)
        self.progress_slider.setRange(0, 0)
        self.progress_slider.sliderMoved.connect(self.set_position)
        self.progress_slider.setStyleSheet("""
            QSlider::groove:horizontal {
                height: 8px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #B1B1B1, stop:1 #c4c4c4);
                margin: 2px 0;
            }
            QSlider::handle:horizontal {
                background: white;
                border: 1px solid #5c5c5c;
                width: 18px;
                margin: -2px 0; /* handles are a little larger than the groove */
                border-radius: 3px;
            }
        """)
        self.current_time_label = QLabel("00:00 / 00:00")
        self.current_time_label.setFont(QFont('Times New Roman', 12))
        self.current_time_label.setStyleSheet("color: black;")
        # 播放时长控制
        self.duration_slider = QSlider(Qt.Horizontal)
        self.duration_slider.setRange(1, 30)  # 默认最大为30秒
        self.duration_slider.setValue(10)     # 默认播放10秒
        self.duration_slider.valueChanged.connect(self.update_duration)
        self.duration_slider.setStyleSheet("""
            QSlider::groove:horizontal {
                height: 8px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #B1B1B1, stop:1 #c4c4c4);
                margin: 2px 0;
            }
            QSlider::handle:horizontal {
                background: white;
                border: 1px solid #5c5c5c;
                width: 18px;
                margin: -2px 0; /* handles are a little larger than the groove */
                border-radius: 3px;
            }
        """)
        # 标签显示当前时间点
        self.time_label = QLabel(f"当前播放间隔：{self.duration_slider.value()} 秒")
        self.time_label.setFont(QFont('宋体', 11))
        self.time_label.setStyleSheet("color: black;")
        # 布局组装
        control_layout = QHBoxLayout()
        control_layout.addWidget(self.current_time_label)
        control_layout.addWidget(self.progress_slider)
        control_widget = QWidget()
        control_widget.setStyleSheet("""
            background-color: #d0d9ff;
        """)
        control_widget.setMinimumHeight(60)  # 设置最小高度
        control_widget.setMaximumHeight(70)
        control_widget.setLayout(control_layout)
        speed_jump_layout = QHBoxLayout()
        speed_jump_layout.addWidget(self.play_button)
        speed_jump_layout.addWidget(self.replay_button)
        speed_jump_layout.addWidget(self.jump_inputmin)
        speed_jump_layout.addWidget(self.jump_inputlabel)
        speed_jump_layout.addWidget(self.jump_input)
        speed_jump_layout.addWidget(self.jump_button)
        speed_jump_layout.addStretch()
        speed_jump_layout.addWidget(self.speed_label)
        speed_jump_layout.addWidget(self.speed_combo)
        duration_layout = QHBoxLayout()
        duration_layout.addWidget(self.time_label)
        duration_layout.addWidget(self.duration_slider)
        self.file_name_label = QLabel("当前文件：未选择")
        self.file_name_label.setFont(QFont('Arial', 8))
        self.file_name_label.setStyleSheet("color: black; padding: 5px;")
        self.file_name_label.setMaximumHeight(35)
        self.file_name_label.setMinimumHeight(30)
        self.file_name_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        # 修改布局顺序 - 将文件名标签放在视频区域上方
        content_layout.insertWidget(0, self.file_name_label)  # 插入到布局最顶部
        content_layout.addWidget(self.video_widget)
        content_layout.addWidget(control_widget)
        content_layout.addLayout(speed_jump_layout)
        content_layout.addLayout(duration_layout)
        # 主布局
        layout.addWidget(file_list_container)
        layout.addLayout(content_layout)
        # 初始化组件
        self._init_components()
    def _init_components(self):
        # 初始化播放器组件
        self.speed_combo.addItems(["0.25x","0.5x", '0.75x',"1x",'1.25x','1.5x',"2x"])
        self.speed_combo.setCurrentIndex(3)
        self.duration_slider.setRange(1, 30)
        self.duration_slider.setValue(10)
        self.progress_slider.setRange(0, 0)
        # 信号连接
        self.open_button.clicked.connect(self.open_file)
        self.file_list.itemDoubleClicked.connect(self.load_video_from_list)
        self.play_button.clicked.connect(self.play_segment)
        self.replay_button.clicked.connect(self.replay_segment)
        self.jump_button.clicked.connect(self.jump_to_time)
        self.duration_slider.valueChanged.connect(self.update_duration)
        self.speed_combo.currentIndexChanged.connect(self.change_speed)
        self.progress_slider.sliderMoved.connect(self.set_position)
        self.media_player.positionChanged.connect(self.update_progress_bar)
        self.media_player.durationChanged.connect(self.update_duration_bar)
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_progress)
    def play_segment(self):
        if not self.media_player.isAvailable():
            return
        if self.media_player.state() == QMediaPlayer.PlayingState:
            self.media_player.pause()
        else:
            self.media_player.setPosition(self.current_position * 1000)
            self.media_player.play()
            self.timer.start(100)  # 启动定时器每100毫秒更新一次
    def replay_segment(self):
        if not self.media_player.isAvailable():
            return
        self.current_position -= self.segment_duration
        if self.current_position < 0:
            self.current_position = 0
        self.media_player.setPosition(self.current_position * 1000)
        self.media_player.play()
        self.timer.start(100)  # 启动定时器每100毫秒更新一次
    def jump_to_time(self):
        try:
            # 获取用户输入的分钟和秒
            minutes_text = self.jump_inputmin.text()
            seconds_text = self.jump_input.text()
            if not minutes_text or not seconds_text:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("无效输入")
                msg_box.setText("请输入有效的数字！")
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setStyleSheet("QMessageBox { background-color: #ffffff; }")
                msg_box.exec_()
                return
            minutes = int(minutes_text)
            seconds = int(seconds_text)
            if minutes < 0 or minutes > 59 or seconds < 0 or seconds > 59:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("无效输入")
                msg_box.setText("分钟和秒必须在 0~59 之间！")
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setStyleSheet("QMessageBox { background-color: #ffffff; }")
                msg_box.exec_()
                return
            # 计算总秒数
            time_point = minutes * 60 + seconds
            # 获取视频总时长（秒）
            total_duration = self.media_player.duration() // 1000
            # 判断是否超过总时长
            if time_point > total_duration:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("无效时间")
                msg_box.setText("跳转时间超过视频总时长")
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setStyleSheet("QMessageBox { background-color: #ffffff; }")
                msg_box.exec_()
                return
            # 设置当前播放位置
            self.current_position = time_point
            self.media_player.setPosition(time_point * 1000)
        except ValueError:
            QMessageBox.warning(self, "无效输入", "请输入有效的数字！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"跳转失败: {str(e)}")
    def update_duration(self, value):
        self.segment_duration = value
        self.time_label.setText(f"当前播放时长: {value} 秒")
    def change_speed(self, index):
        speed_str = self.speed_combo.itemText(index)
        self.current_speed = float(speed_str[:-1])
        self.media_player.setPlaybackRate(self.current_speed)
    def open_file(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(self, "打开视频文件", "", "AVI Files (*.avi);;All Files (*)", options=options)
        if files:
            self.file_list.addItems(files)
            self.current_file_path = files[0]  # 记录第一个文件路径
            self.update_file_name_label()
    def load_video_from_list(self, item):
        file_path = item.text()
        self.current_file_path = file_path     # 更新当前文件路径
        self.update_file_name_label()
        self.load_video(file_path)
    def load_video(self, file_path):
        if file_path:
            self.media_player.setMedia(QMediaContent(QUrl.fromLocalFile(file_path)))
            self.current_position = 0
            self.media_player.setPosition(0)
            self.media_player.play()
            self.media_player.pause()  # 立即暂停以显示第一帧
            self.timer.start(100)  # 启动定时器每100毫秒更新一次
            if file_path != self.current_file_path:
                self.current_file_path = file_path
                self.update_file_name_label()
    def update_file_name_label(self):
        """更新文件名显示，显示完整文件名"""
        if self.current_file_path:
            file_name = self.current_file_path.split('/')[-1]  # 提取文件名部分
            self.file_name_label.setText(f"当前文件：{file_name}")
        else:
            self.file_name_label.setText("当前文件：未选择")
    def set_position(self, position):
        self.media_player.setPosition(position)
    def update_progress(self):
        position = self.media_player.position()
        current_time = self.format_time(position)
        total_time = self.format_time(self.media_player.duration())
        self.current_time_label.setText(f"{current_time} / {total_time}")
        if position >= (self.current_position + self.segment_duration) * 1000:
            self.media_player.pause()
            self.current_position += self.segment_duration
    def update_progress_bar(self, position):
        self.progress_slider.setValue(position)
    def update_duration_bar(self, duration):
        self.progress_slider.setRange(0, duration)
    def format_time(self, milliseconds):
        seconds = int(milliseconds / 1000)
        minutes = int(seconds / 60)
        seconds %= 60
        return f"{minutes:02}:{seconds:02}"
class DataEncoderModule(QWidget):
    def __init__(self, group_name=""):
        super().__init__()
        self.group_name = group_name
        self.modules = []
        self.selected_data = []
        self.history = []
        self.initUI()
    def initUI(self):
        # 主布局
        main_layout = QHBoxLayout(self)
        # 左侧区域
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        # 顶部输入区域
        top_widget = QWidget()
        top_layout = QHBoxLayout()
        self.comment_entry = QLineEdit()
        self.comment_entry.setPlaceholderText("输入注释...")
        self.module_edit = QLabel()
        self.module_edit.setText("维度编辑")
        self.module_edit.setFont(QFont('黑体', 12))
        self.module_edit.setStyleSheet("font-size: 16px;")
        self.add_module_btn = QToolButton()
        self.add_module_btn.setText('+')
        self.add_module_btn.setFixedSize(30, 30)
        self.add_module_btn.setStyleSheet("font-size: 16px; background-color: #a6baff;")
        self.remove_module_btn = QToolButton()
        self.remove_module_btn.setText('-')
        self.remove_module_btn.setFixedSize(30, 30)
        self.remove_module_btn.setStyleSheet("font-size: 16px; background-color: #a6baff;")
        top_layout.addWidget(self.comment_entry)
        top_layout.addWidget(self.module_edit)
        top_layout.addWidget(self.add_module_btn)
        top_layout.addWidget(self.remove_module_btn)
        top_widget.setLayout(top_layout)
        # 模块区域滚动区
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.module_container = QWidget()
        self.module_layout = QVBoxLayout()
        self.module_layout.setAlignment(Qt.AlignTop)
        self.module_container.setAutoFillBackground(True)
        p = self.module_container.palette()
        p.setColor(self.module_container.backgroundRole(), QColor(166, 186, 255))
        self.module_container.setPalette(p)
        self.module_container.setLayout(self.module_layout)
        self.scroll_area.setWidget(self.module_container)
        # 按钮区域
        btn_widget = QWidget()
        btn_layout = QHBoxLayout()
        self.confirm_btn = QPushButton('确认')
        self.confirm_btn.setFont(QFont('黑体', 10))
        self.confirm_btn.setFixedSize(60, 30)
        self.continue_btn = QPushButton('继续')
        self.continue_btn.setFont(QFont('黑体', 10))
        self.continue_btn.setFixedSize(60, 30)
        self.undo_btn = QPushButton('回退')
        self.undo_btn.setFont(QFont('黑体', 10))
        self.undo_btn.setFixedSize(60, 30)
        self.clear_btn = QPushButton('清空')
        self.clear_btn.setFont(QFont('黑体', 10))
        self.clear_btn.setFixedSize(60, 30)
        self.output_btn = QPushButton('输出')
        self.output_btn.setFont(QFont('黑体', 10))
        self.output_btn.setFixedSize(60, 30)
        btn_layout.addWidget(self.confirm_btn)
        btn_layout.addWidget(self.continue_btn)
        btn_layout.addWidget(self.undo_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.output_btn)
        btn_widget.setLayout(btn_layout)
        # 右侧结果区域
        result_widget = QWidget()
        result_layout = QVBoxLayout()
        # 时间单位设置
        time_widget = QWidget()
        time_layout = QHBoxLayout()
        self.time_unit_label = QLabel("时间单位:")
        self.time_unit_label.setAlignment(Qt.AlignVCenter|Qt.AlignLeft)
        self.time_unit_label.setFont(QFont('宋体', 11))
        self.time_unit_entry = QLineEdit()
        self.time_unit_entry.setPlaceholderText("秒")
        self.time_unit_entry.setFont(QFont('Times New Roman', 8))
        self.time_unit_entry.setFixedSize(30,30)
        self.time_unit_entry.setValidator(QIntValidator(0, 30, self)) 
        self.timestartlabel = QLabel("起始时间:")
        self.timestartlabel.setAlignment(Qt.AlignVCenter|Qt.AlignLeft)
        self.timestartlabel.setFont(QFont('宋体', 11))
        self.startmin = QLineEdit()
        self.startmin.setFont(QFont('Times New Roman', 8))
        self.startmin.setPlaceholderText("分")
        self.startmin.setFixedSize(30, 30)
        self.startmin.setValidator(QIntValidator(0, 59, self))  # 添加验证器
        self.starttimelabel = QLabel(':')
        self.starttimelabel.setFont(QFont('Times New Roman', 11))
        self.starttimelabel.setStyleSheet("color: black;")
        self.starttimelabel.setFixedSize(5, 30)
        self.startsec = QLineEdit()
        self.startsec.setFont(QFont('Times New Roman', 8))
        self.startsec.setPlaceholderText("秒")
        self.startsec.setFixedSize(30, 30)
        self.startsec.setValidator(QIntValidator(0, 59, self)) 
        #self.clearhistory_btn = QPushButton('清空记录')
        #self.clearhistory_btn.setFont(QFont('黑体', 10))
        #self.clearhistory_btn.setFixedSize(75, 35)
        self.row_unit_label = QLabel("选定行号:")
        self.row_unit_label.setAlignment(Qt.AlignVCenter|Qt.AlignLeft)
        self.row_unit_label.setFont(QFont('宋体', 11))
        self.row_number_entry= QLineEdit()
        self.row_number_entry.setPlaceholderText("1")
        self.row_number_entry.setFixedSize(30,30)
        self.row_number_entry.setFont(QFont('Times New Roman', 10))
        self.insert_btn = QPushButton('插入行')
        self.insert_btn.setFont(QFont('黑体', 10))
        self.insert_btn.setFixedSize(70, 35)
        self.delete_btn = QPushButton('删除行')
        self.delete_btn.setFont(QFont('黑体', 10))
        self.delete_btn.setFixedSize(70, 35)
        time_layout.addWidget(self.time_unit_label)
        time_layout.addWidget(self.time_unit_entry)
        time_layout.addWidget(self.timestartlabel)
        time_layout.addWidget(self.startmin)
        time_layout.addWidget(self.starttimelabel)
        time_layout.addWidget(self.startsec)
        time_layout.addStretch()
        #time_layout.addWidget(self.clearhistory_btn)
        time_layout.addWidget(self.row_unit_label)
        time_layout.addWidget(self.row_number_entry)
        time_layout.addWidget(self.insert_btn)
        time_layout.addWidget(self.delete_btn)
        time_widget.setAutoFillBackground(True)
        p = top_widget.palette()
        p.setColor(time_widget.backgroundRole(), QColor(208, 217, 255))
        time_widget.setPalette(p)
        time_widget.setLayout(time_layout)
        # 表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(5)
        self.result_table.setHorizontalHeaderLabels(["序号", "时间", "维度1", "维度2", "注释"])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        result_layout.addWidget(time_widget)
        result_layout.addWidget(self.result_table)
        result_widget.setLayout(result_layout)
        # 组装左侧布局
        left_layout.addWidget(top_widget)
        left_layout.addWidget(self.scroll_area)
        left_layout.addWidget(btn_widget)
        # 主布局
        main_layout.addLayout(left_layout, 1)
        main_layout.addWidget(result_widget, 3)
        # 初始化信号连接
        self.add_module_btn.clicked.connect(self.add_module)
        self.remove_module_btn.clicked.connect(self.remove_module)
        self.confirm_btn.clicked.connect(self.confirm)
        self.continue_btn.clicked.connect(self.continue_)
        self.undo_btn.clicked.connect(self.undo)
        self.clear_btn.clicked.connect(self.clear)
        self.output_btn.clicked.connect(self.output)
        self.result_table.itemDoubleClicked.connect(self.on_cell_double_clicked)
        #self.clearhistory_btn.clicked.connect(self.clear_history)
        self.insert_btn.clicked.connect(self.insert_row_at_input)
        self.delete_btn.clicked.connect(self.delete_row_at_input)
        # 初始化模块
        self.time_unit_entry.textChanged.connect(self.validate_time_unit)
        self.time_unit_entry.textChanged.connect(self.update_table)
        self.row_number_entry.textChanged.connect(self.get_valid_row_index)
        self.startmin.textChanged.connect(self.validate_input)
        self.startsec.textChanged.connect(self.validate_input)
        self.continue_btn.setEnabled(False)
        self.undo_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.output_btn.setEnabled(False)
        self.insert_btn.setEnabled(False)
        self.delete_btn.setEnabled(False)
        self.row_number_entry.setEnabled(False)
    def validate_time_unit(self):
        timeunit=self.time_unit_entry.text()
        if timeunit.isdigit():
                val = int(timeunit)
                if 0 <= val <= 30:
                    self.time_unit_entry.setStyleSheet("border: 1px solid black;")
                else:
                    self.time_unit_entry.setStyleSheet("border: 1px solid red;")
        else:
            self.time_unit_entry.setStyleSheet("border: 1px solid red;")
    def validate_input(self):
        for line_edit in [self.startmin, self.startsec]:
            text = line_edit.text()
            if text.isdigit():
                val = int(text)
                if 0 <= val <= 59:
                    line_edit.setStyleSheet("border: 1px solid black;")
                else:
                    line_edit.setStyleSheet("border: 1px solid red;")
            else:
                line_edit.setStyleSheet("border: 1px solid red;")
    def add_module(self):
        module_number = len(self.modules) + 1
        module = Module(self)
        self.modules.append(module)
        self.module_layout.addWidget(module)
        module.removed.connect(self.check_module_count)
        self.check_module_count()
    def check_module_count(self):
        if len(self.modules) > 0:
            self.remove_module_btn.setEnabled(True)
        else:
            self.remove_module_btn.setEnabled(False)
    def remove_module(self):
        if self.modules:
            module = self.modules.pop()
            module.setParent(None)
            module.deleteLater()
            self.check_module_count()
            if self.confirm_btn.isEnabled():
                self.update_table_columns()
    def confirm(self):
        self.confirm_btn.setEnabled(False)
        self.clear_btn.setEnabled(True)
        self.add_module_btn.setEnabled(False)
        self.remove_module_btn.setEnabled(False) 
        self.continue_btn.setEnabled(True)
        self.output_btn.setEnabled(True)
        self.undo_btn.setEnabled(False)
        self.time_unit_entry.setEnabled(False)
        self.startmin.setEnabled(False)
        self.startsec.setEnabled(False)
        self.insert_btn.setEnabled(True)
        self.delete_btn.setEnabled(True)
        self.row_number_entry.setEnabled(True)
        self.selected_data = []
        self.history = [copy.deepcopy(self.selected_data)]
        self.update_table_columns()
        self.update_table()
    def continue_(self):
        current_row = []
        for module in self.modules:
            selected = module.get_selected_entry()
            current_row.append(selected)
        current_row.append(self.comment_entry.text())
        self.selected_data.append(current_row)
        self.history.append(copy.deepcopy(self.selected_data))
        self.update_undo_state()
        self.update_table()
        self.comment_entry.clear()
    def undo(self):
        if len(self.history) >= 2:
            self.history.pop()
            self.selected_data = copy.deepcopy(self.history[-1])
            self.update_table()
            self.update_undo_state()
            self.comment_entry.clear()
    def update_undo_state(self):
        self.undo_btn.setEnabled(len(self.history) > 1)
    def update_table_columns(self):
        columns = ["序号", "时间"] + [module.name for module in self.modules] + ["注释"]
        self.result_table.setColumnCount(len(columns))
        self.result_table.setHorizontalHeaderLabels(columns)
    def format_seconds(self,seconds):
        hours = seconds // 3600
        remaining = seconds % 3600
        minutes = remaining // 60
        secs = remaining % 60
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{secs:02d}"
        else:
            return f"{minutes:02d}:{secs:02d}"
    def format_seconds2(seconds):
        hours, remainder = divmod(seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        return f"{hours}:{minutes:02d}:{secs:02d}"
    def update_table(self):
        self.result_table.setRowCount(0)
        try:
            starttime=int(self.startmin.text())*60+int(self.startsec.text())
        except:
            starttime=0
        try:
            time_unit = int(self.time_unit_entry.text())
        except:
            time_unit = 1
        for idx, row in enumerate(self.selected_data, 1):
            self.result_table.insertRow(idx-1)
            time_value = idx * time_unit
            time_prevalue= starttime+(idx - 1)*time_unit
            timelength= self.format_seconds(time_prevalue)# Format seconds to HH:MM:SS
            values = [idx, timelength] + row
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                self.result_table.setItem(idx-1, col, item)
        self.adjust_column_widths()
    def adjust_column_widths(self):
        for col in range(self.result_table.columnCount()):
            max_width = 100
            for row in range(self.result_table.rowCount()):
                item = self.result_table.item(row, col)
                if item:
                    width = len(item.text()) * 10
                    if width > max_width:
                        max_width = width
            self.result_table.setColumnWidth(col, max_width if max_width < 200 else 200)
    def sync_table_to_data(self):
        """将表格当前内容同步到self.selected_data"""
        self.selected_data = []
        for row in range(self.result_table.rowCount()):
            row_data = []
            for col in range(self.result_table.columnCount()):
                item = self.result_table.item(row, col)
                row_data.append(item.text() if item else "")
            self.selected_data.append(row_data)
    
    def output(self):
        """输出表格内容到Excel文件"""
        # 强制提交未提交的编辑内容
        self.result_table.clearSelection()

        # 同步表格内容到 self.selected_data
        self.sync_table_to_data()

        # 获取表格列名
        columns = [self.result_table.horizontalHeaderItem(col).text() 
                for col in range(1,self.result_table.columnCount())]

        # 构建 DataFrame
        if not self.selected_data:
            df = pd.DataFrame(columns=columns)
        else:
            filtered_data = [row[1:] for row in self.selected_data]
            df = pd.DataFrame(filtered_data, columns=columns)

        # 选择保存路径
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存数据", "", "Excel 文件 (*.xlsx);;所有文件 (*)"
        )
        if not file_path:
            return

        # 导出到 Excel
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='数据', index=False)

                worksheet = writer.sheets['数据']
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    worksheet.column_dimensions[column].width = max_length + 4
                    for cell in col:
                        cell.alignment = Alignment(horizontal='center')

                for cell in worksheet["1:1"]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')

            QMessageBox.information(self, "成功", "数据已成功导出！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")
    def clear(self):
        self.selected_data = []
        self.history = [copy.deepcopy(self.selected_data)]
        self.update_undo_state()
        self.result_table.setRowCount(0)
        #for module in self.modules:
            #module.reset()
        self.comment_entry.clear()
        self.add_module_btn.setEnabled(True)
        self.remove_module_btn.setEnabled(True)
        self.undo_btn.setEnabled(False)
        self.continue_btn.setEnabled(False)
        self.output_btn.setEnabled(False)
        self.confirm_btn.setEnabled(True)
        self.time_unit_entry.setEnabled(True)
        self.startmin.setEnabled(True)
        self.startsec.setEnabled(True)
    def on_cell_double_clicked(self, item):
        row = item.row()
        col = item.column()
        if col == len(self.modules) + 2:  # 注释列
            entry = QLineEdit(item.text())
            entry.editingFinished.connect(lambda: self.save_cell(entry, row, col))
            self.result_table.setCellWidget(row, col, entry)
            entry.setFocus()
    def save_cell(self, entry, row, col):
        new_value = entry.text()
        item = self.result_table.item(row, col)
        if item:
            item.setText(new_value)
        self.result_table.removeCellWidget(row, col)
        # 更新内部数据
        if row < len(self.selected_data):
            self.selected_data[row][col-2] = new_value
    def update_table_with_data(self, data):
        self.result_table.setRowCount(len(data))
        for row, items in enumerate(data):
            for col, item in enumerate(items):
                self.result_table.setItem(row, col, QTableWidgetItem(str(item)))
    def validate_number(self, value):
        if value == "":
            return True
        try:
            int(value)
            return True
        except:
            return False
    def get_valid_row_index(self):
        row_input = self.row_number_entry.text().strip()
        if not row_input:
            return None
        try:
            row_number = int(row_input)
            if 1 <= row_number <= len(self.selected_data):
                return row_number - 1  # 转换为 0-based 索引
            else:
                raise ValueError("Row number out of range")
        except ValueError:
            # 可选：清空无效输入
            self.row_number_entry.clear()
            return None
    def insert_row_at_input(self):
        index = self.get_valid_row_index()
        if index is None:
            return
        current_row = []
        for module in self.modules:
            selected = module.get_selected_entry()
            current_row.append(selected)
        current_row.append(self.comment_entry.text().strip())
        self.selected_data.insert(index, current_row)
        self.history.append(copy.deepcopy(self.selected_data))
        self.update_undo_state()
        self.update_table()
        self.comment_entry.clear()
    def delete_row_at_input(self):
        index = self.get_valid_row_index()
        if index is None:
            return
        if 0 <= index < len(self.selected_data):
            self.selected_data.pop(index)
            self.history.append(copy.deepcopy(self.selected_data))
            self.update_undo_state()
            self.update_table()
class Module(QWidget):
    removed = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__()
        self.parent = parent
        self.name = f"维度{len(parent.modules)+1 if parent else 1}"
        self.blocks = []
        self.selected_index = -1
        self.initUI()
    def initUI(self):
        # 整体布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        # 模块头部
        header_frame = QFrame()
        header_frame.setStyleSheet("background-color: #ffffff;")
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(5, 0, 5, 0)
        header_layout.setSpacing(0)
        self.label = QLabel(self.name)
        self.label.setStyleSheet("font-size: 12pt; font-family: 黑体;background-color: #d0d9ff;")
        self.label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        self.label.setMinimumWidth(100)
        # 添加可编辑的输入框
        self.edit_entry = QLineEdit()
        self.edit_entry.setStyleSheet("background-color: white;")
        self.edit_entry.setMaximumWidth(100)
        self.edit_entry.setFont(QFont("黑体", 12))
        self.edit_entry.setVisible(False)
        self.edit_entry.editingFinished.connect(self.save_label)
        self.add_block_btn = QToolButton()
        self.add_block_btn.setText('+')
        self.add_block_btn.setFixedSize(30, 30)
        self.add_block_btn.setStyleSheet("font-size: 12px; background-color: #d0d9ff;")
        self.remove_block_btn = QToolButton()
        self.remove_block_btn.setText('-')
        self.remove_block_btn.setFixedSize(30, 30)
        self.remove_block_btn.setStyleSheet("font-size: 12px; background-color: #d0d9ff;")
        # 创建下拉箭头标志（替换原来的按钮）
        self.dropdown_indicator = QLabel(self)
        imagepath=os.path.join(os.path.dirname(__file__), 'downarrow(1).png')
        pixmap = QPixmap(imagepath)
        self.dropdown_indicator.setPixmap(pixmap.scaled(30, 20)) 
        self.dropdown_indicator.setStyleSheet("background-color: white;")
        self.dropdown_indicator.setAlignment(Qt.AlignBottom)
          # Center the icon in the label
        # 修改布局顺序（箭头放在stretch中间位置）
        header_layout.addWidget(self.label)
        header_layout.addWidget(self.edit_entry)
        header_layout.addStretch()  # 添加可伸缩空间
        header_layout.addWidget(self.dropdown_indicator)  # 添加箭头标志
        header_layout.addStretch()  # 添加对称伸缩空间保持居中
        header_layout.addWidget(self.add_block_btn)
        header_layout.addWidget(self.remove_block_btn)
        header_frame.setLayout(header_layout)
        # 块容器
        self.block_container = QWidget()
        self.block_layout = QGridLayout()
        self.block_layout.setContentsMargins(5, 5, 5, 5)
        self.block_layout.setSpacing(5)
        self.block_container.setLayout(self.block_layout)
        # 事件处理
        header_frame.mousePressEvent = self.handle_header_click
        header_frame.mouseDoubleClickEvent = self.edit_label
        # 在 header_frame 初始化后添加以下代码
        # 添加到布局
        layout.addWidget(header_frame)
        layout.addWidget(self.block_container)
        # 信号连接
        self.add_block_btn.clicked.connect(self.add_block)
        self.remove_block_btn.clicked.connect(self.remove_block)
    def edit_label(self, event=None):
        """双击编辑模块名称"""
        self.label.setVisible(False)
        self.edit_entry.setText(self.label.text())
        self.edit_entry.setVisible(True)
        self.edit_entry.setFocus()
    def save_label(self):
        """保存模块名称"""
        new_text = self.edit_entry.text()
        if new_text:
            self.label.setText(new_text)
            self.name = new_text
            self.edit_entry.setVisible(False)
            self.label.setVisible(True)
            # 如果处于确认状态，更新表格列名
            if self.parent and self.parent.confirm_btn.isEnabled() == False:
                if hasattr(self.parent, 'update_table_columns'):
                    self.parent.update_table_columns()
                    self.parent.update_table()
    def add_block(self):
        """添加新块"""
        index = len(self.blocks)
        row = index // 2
        col = index % 2
        # 创建容器
        block_frame = QFrame()
        block_frame.setStyleSheet("background-color: #e7e9fd;")
        block_layout = QHBoxLayout()
        block_layout.setContentsMargins(0, 0, 0, 0)
        block_layout.setSpacing(5)
        # 输入框
        entry = QLineEdit()
        entry.setStyleSheet("background-color: white;")
        entry.setFont(QFont("Times New Roman", 11))
        entry.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        # 单选框
        checkbox = QCheckBox()
        checkbox.setStyleSheet("margin-left: 5px;")
        checkbox.clicked.connect(lambda: self.select_block(index))
        # 添加到布局
        block_layout.addWidget(entry)
        block_layout.addWidget(checkbox)
        block_frame.setLayout(block_layout)
        # 布局管理
        self.block_layout.addWidget(block_frame, row, col)
        self.blocks.append({'entry': entry, 'checkbox': checkbox, 'frame': block_frame})
    def remove_block(self):
        """删除最后一个块"""
        if self.blocks:
            block = self.blocks.pop()
            block['frame'].deleteLater()
            self.rearrange_blocks()
    def rearrange_blocks(self):
        """重新排列剩余的块"""
        for i, block in enumerate(self.blocks):
            row = i // 2
            col = i % 2
            self.block_layout.addWidget(block['frame'], row, col)
    def select_block(self, index):
        """选择块"""
        if self.selected_index == index:
            self.selected_index = -1
            self.blocks[index]['checkbox'].setChecked(False)
        else:
            if self.selected_index != -1 and self.selected_index < len(self.blocks):
                self.blocks[self.selected_index]['checkbox'].setChecked(False)
            self.selected_index = index
            self.blocks[index]['checkbox'].setChecked(True)
    def get_selected_entry(self):
        """获取选中块的文本"""
        if self.selected_index != -1 and self.selected_index < len(self.blocks):
            return self.blocks[self.selected_index]['entry'].text()
        return ""
    def reset(self):
        """重置模块"""
        self.selected_index = -1
        for block in self.blocks:
            block['entry'].clear()
            block['checkbox'].setChecked(False)
    def mouseDoubleClickEvent(self, event):
        """双击编辑模块名称"""
        self.edit_label()
    def handle_header_click(self, event):
        # 获取点击位置
        pos = event.pos()
        # 获取按钮的几何区域
        add_btn_rect = self.add_block_btn.geometry()
        remove_btn_rect = self.remove_block_btn.geometry()
        # 如果点击位置在按钮区域内，则不执行 toggle
        if add_btn_rect.contains(pos) or remove_btn_rect.contains(pos):
            return
        # 否则调用 toggle_blocks
        self.toggle_blocks()
    def toggle_blocks(self):
        if self.block_container.isVisible():
            self.block_container.hide()
        else:
            self.block_container.show()
class ClusterAnalysisModule(QWidget):
    def __init__(self, tabmodule=None):
        super().__init__()
        self.dimensions = []  # 存储编码维度
        self.student_mincount = 2  # 默认学生人数
        self.student_maxcount = 3
        self.group_count = 22  # 默认小组数量
        self.n_clusters = 3  # 设置聚类数目
        self.tabmodule = tabmodule 
        self.plot_edit=False
        self.initUI()
    def initUI(self):
        # 创建主布局
        main_layout = QHBoxLayout(self)
        # 左侧面板 - 参数设置
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        # 小组和学生人数设置
        group_student_layout = QHBoxLayout()
        self.group_label = QLabel("小组数量:")
        self.group_label.setFont(QFont('黑体', 10))
        self.group_label.setMinimumHeight(35)
        self.group_spin = QSpinBox()
        self.group_spin.setMinimumHeight(35)
        self.group_spin.setValue(self.group_count)
        self.group_spin.valueChanged.connect(self.update_group_count)
        self.maxstudent_label = QLabel("最多人数:")
        self.maxstudent_label.setFont(QFont('黑体', 10))
        self.maxstudent_label.setMinimumHeight(35)
        self.maxstudent_spin = QSpinBox()
        self.maxstudent_spin.setMinimumHeight(35)
        self.maxstudent_spin.setValue(self.student_maxcount)
        self.maxstudent_spin.valueChanged.connect(self.update_student_maxcount)
        self.minstudent_label = QLabel("最少人数:")
        self.minstudent_label.setFont(QFont('黑体', 10))
        self.minstudent_label.setMinimumHeight(35)
        self.minstudent_spin = QSpinBox()
        self.minstudent_spin.setMinimumHeight(35)
        self.minstudent_spin.setValue(self.student_mincount)
        self.minstudent_spin.valueChanged.connect(self.update_student_mincount)
        self.n_clusters_label = QLabel("聚类数量:")
        self.n_clusters_label.setFont(QFont('黑体', 10))
        self.n_clusters_label.setMinimumHeight(35)
        self.n_clusters_spin = QSpinBox()   
        self.n_clusters_spin.setMinimumHeight(35)
        self.n_clusters_spin.setValue(self.n_clusters)
        self.n_clusters_spin.valueChanged.connect(self.update_n_clusters)
        group_student_layout.addWidget(self.group_label)
        group_student_layout.addWidget(self.group_spin)
        group_student_layout.addWidget(self.minstudent_label)
        group_student_layout.addWidget(self.minstudent_spin)
        group_student_layout.addWidget(self.maxstudent_label)
        group_student_layout.addWidget(self.maxstudent_spin)
        group_student_layout.addWidget(self.n_clusters_label)
        group_student_layout.addWidget(self.n_clusters_spin)
        # 维度管理
        self.copy_encoding_btn = QPushButton('传递编码')
        self.copy_encoding_btn.setFont(QFont('黑体', 10))
        self.copy_encoding_btn.setFixedHeight(35)
        self.copy_encoding_btn.setFixedWidth(100)
        self.copy_encoding_btn.clicked.connect(self.copy_encoding_from_other_tab)
        dimension_header_layout = QHBoxLayout()
        self.dimension_label = QLabel("编码维度:")
        self.dimension_label.setFont(QFont('黑体', 12))
        self.dimension_label.setFixedHeight(35)
        self.add_dimension_btn = QPushButton('+')
        self.add_dimension_btn.clicked.connect(self.add_dimension)
        self.remove_dimension_btn = QPushButton('-')
        self.remove_dimension_btn.clicked.connect(self.remove_dimension)
        dimension_header_layout.addWidget(self.copy_encoding_btn)
        dimension_header_layout.addStretch()
        dimension_header_layout.addWidget(self.dimension_label)
        dimension_header_layout.addWidget(self.add_dimension_btn)
        dimension_header_layout.addWidget(self.remove_dimension_btn)
        # 维度容器
        self.dimension_container = QWidget()
        self.dimension_container.setStyleSheet("background-color: #d0d9ff;")  # 添加背景色
        self.dimension_layout = QVBoxLayout()
        self.dimension_layout.setAlignment(Qt.AlignTop)
        self.dimension_container.setLayout(self.dimension_layout)
        # 文件选择按钮
        self.file_select_btn = QPushButton('选择文件')
        self.file_select_btn.setFont(QFont('黑体', 10))
        self.file_select_btn.setMinimumHeight(35)
        self.file_select_btn.clicked.connect(self.select_files)
        # 分析按钮
        self.analyze_btn = QPushButton('开始分析')
        self.analyze_btn.setFont(QFont('黑体', 10))
        self.analyze_btn.setMinimumHeight(35)
        self.analyze_btn.clicked.connect(self.perform_analysis)
        analysis_layout = QHBoxLayout()
        analysis_layout.addWidget(self.file_select_btn)
        analysis_layout.addWidget(self.analyze_btn)
        # 组装左侧面板
        left_layout.addLayout(group_student_layout)
        left_layout.addStretch()
        left_layout.addLayout(dimension_header_layout)
        left_layout.addWidget(self.dimension_container)
        left_layout.addStretch()
        left_layout.addLayout(analysis_layout)
        left_layout.addStretch()
        left_panel.setLayout(left_layout)
        # 右侧面板 - 结果展示
        right_panel = QWidget()
        right_layout = QHBoxLayout()
        # 结果表格
        self.result_label = QLabel("数据预览")
        self.result_label.setFont(QFont('黑体', 12))
        self.result_label.setStyleSheet("background-color: #a6baff;")
        self.result_label.setAlignment(Qt.AlignCenter)
        self.result_label.setFixedHeight(35)
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(5)
        self.result_table.setHorizontalHeaderLabels(["组别", "时间", "维度1", "维度2", "维度3"])
        self.notice = QLabel("请注意：组别序号按文件导入顺序自动分配")
        self.notice.setFont(QFont('黑体', 8))
        self.table_container = QWidget()
        table_layout = QVBoxLayout(self.table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.addWidget(self.result_label)
        table_layout.addWidget(self.result_table)
        table_layout.addWidget(self.notice)
        # 图表区域（新增）
        self.canvas_label = QLabel("聚类结果")
        self.canvas_label.setFont(QFont('黑体', 12))
        self.canvas_label.setFixedHeight(35)
        self.figure = Figure(figsize=(15, 10))
        self.canvas = FigureCanvas(self.figure)
        self.canvas_btn = QPushButton('图片编辑和导出')
        self.canvas_btn.setFont(QFont('黑体', 10))
        self.canvas_btn.setMinimumHeight(35)
        self.canvas_btn.clicked.connect(self.canvas_edit)
        self.canvas_container = QWidget()
        self.canvas_container.setStyleSheet("background-color: #d0d9ff;")
        canvas_layout = QVBoxLayout(self.canvas_container)
        canvas_layout.setContentsMargins(10, 10, 10, 10)  # 左、上、右、下边距
        canvas_layout.addWidget(self.canvas_label)
        canvas_layout.addWidget(self.canvas)
        canvas_layout.addWidget(self.canvas_btn)
            # 新增聚类结果显示区域
        cluster_result_group = QGroupBox("聚类结果映射")
        cluster_result_group.setFont(QFont("黑体", 12))
        self.canvas_btn.setMinimumHeight(35)
        self.cluster_result_layout = QVBoxLayout()
        self.cluster_result_layout.setSpacing(5)
        self.cluster_result_layout.setContentsMargins(10, 10, 10, 10)
        # 添加滚动区域支持大量数据
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_content.setLayout(self.cluster_result_layout)
        scroll_area.setWidget(scroll_content)
        # 将滚动区域添加到布局
        left_layout.addWidget(cluster_result_group)
        cluster_result_group.setLayout(self.cluster_result_layout)
        # 组装右侧面板
        right_layout.addWidget(self.table_container, 3)  # 占3份宽度
        right_layout.addWidget(self.canvas_container, 3) 
        right_panel.setLayout(right_layout)
        # 主布局
        main_layout.addWidget(left_panel, 1)
        main_layout.addWidget(right_panel, 3)
        # 初始化默认维度
        self.add_dimension()
        if self.student_mincount > 0:
            self.add_dimension()
        if self.student_mincount > 1:
            self.add_dimension()
    def add_dimension(self):
        dimension_number = len(self.dimensions) + 1
        dimension = Dimension(self, dimension_number)
        self.dimensions.append(dimension)
        self.dimension_layout.addWidget(dimension)
        # 更新维度标签
        for i, dim in enumerate(self.dimensions):
            dim.dimension_number = i + 1
            dim.update_label()
    def remove_dimension(self):
        if self.dimensions:
            dimension = self.dimensions.pop()
            dimension.setParent(None)
            dimension.deleteLater()
            # 更新剩余维度的标签
            for i, dim in enumerate(self.dimensions):
                dim.dimension_number = i + 1
                dim.update_label()
    def update_group_count(self, value):
        self.group_count = value
    def update_student_maxcount(self, value):
        self.student_maxcount = value
    def update_student_mincount(self, value):
        self.student_mincount = value
    def update_n_clusters(self, value):
        self.n_clusters = value
    def get_encoding_columns(self):
        """根据小组人数和维度生成编码列名"""
        columns = []
        for student in range(1, self.student_maxcount + 1):
            for dim in range(1, len(self.dimensions) + 1):
                columns.append(f'S{student}Eoding_{dim}')
        return columns
    def perform_analysis(self):
        # 这里实现你的聚类分析逻辑
        # 执行数据处理和分析
        try:
            # 检查文件是否已选择
            if not hasattr(self, 'selected_files') or not self.selected_files:
                QMessageBox.warning(self, "未选择文件", "请先选择要分析的Excel文件")
                return
            # 获取当前设置值
            group_count = self.group_spin.value()
            n_clusters = self.n_clusters_spin.value()
            minstudent = self.minstudent_spin.value()
            maxstudent = self.maxstudent_spin.value()
            valid_file_count = len(self.selected_files)
            # 条件1：最少人数必须 > 0
            if minstudent <= 0:
                QMessageBox.critical(
                    self, 
                    "配置错误",
                    f"最少人数({minstudent})必须大于0\n"
                    "请调整【最少人数】设置"
                )
                return
            # 条件2：最多人数必须 ≥ 最少人数
            if maxstudent < minstudent:
                QMessageBox.critical(
                    self, 
                    "配置错误",
                    f"最多人数({maxstudent})必须大于或等于最少人数({minstudent})\n"
                    "请调整【最多人数】设置"
                )
                return
            # 条件3：最多人数 ≤ 检测到的 S 后缀最大值
            max_detected = 0
            for file_path in self.selected_files:
                df = pd.read_excel(file_path, nrows=0)  # 仅读取列名
                for col in df.columns:
                    if col.startswith('S') and 'Eoding' in col:
                        try:
                            student_num = int(col[1:].split('Eoding')[0])  # 提取 S 后的数字
                            max_detected = max(max_detected, student_num)
                        except:
                            continue
            if maxstudent > max_detected:
                QMessageBox.critical(
                    self, 
                    "配置错误",
                    f"最多人数({maxstudent})不能超过检测到的学生序号最大值({max_detected})\n"
                    "请调整【最多人数】设置"
                )
                return
             # 条件4：编码维度数量 ≤ 检测到的 Eoding_ 后缀最大值
            max_dim_num = 0
            for file_path in self.selected_files:
                df = pd.read_excel(file_path, nrows=0)  # 仅读取列名
                for col in df.columns:
                    if 'Eoding' in col:
                        try:
                            dim_str = col.split('Eoding')[1]
                            if dim_str.startswith('_') and dim_str[1:].isdigit():
                                dim_num = int(dim_str[1:])  # 提取 _ 后的数字
                                max_dim_num = max(max_dim_num, dim_num)
                        except:
                            continue
            dimension_count = len(self.dimensions)
            if dimension_count > max_dim_num:
                QMessageBox.critical(
                    self, 
                    "配置错误",
                    f"编码维度数量({dimension_count})不能超过检测到的编码维度最大值({max_dim_num})\n"
                    "请调整【编码维度】数量"
                )
                return
            # 获取编码配置
            encoding_config = []
            for dim in self.dimensions:
                codes = dim.get_codes()
                if codes:
                    encoding_config.append(codes)
            # 条件5：必须至少存在一个编码类型
            if not encoding_config:
                QMessageBox.critical(
                    self, 
                    "配置错误",
                    "未定义任何编码类型\n"
                    "请在【编码维度】中至少添加一个编码类型"
                )
                return
            # 条件6：小组数量必须等于有效文件数
            if group_count != valid_file_count:
                QMessageBox.critical(
                    self, 
                    "配置错误",
                    f"小组数量({group_count})必须等于有效文件数({valid_file_count})\n"
                    "请调整【小组数量】设置"
                )
                return
            # 条件7：聚类数量必须小于小组数量
            if n_clusters >= group_count:
                QMessageBox.critical(
                    self, 
                    "配置错误",
                    f"聚类数量({n_clusters})必须小于小组数量({group_count})\n"
                    "请调整【聚类数量】设置"
                )
                return
            # 所有检查通过后执行分析
            self.run_analysis(encoding_config)
        except Exception as e:
            print("分析出错:", str(e))
    def display_results(self, clustered_data=None):
        """在表格中显示分析结果 + 可视化聚类结果"""
        # 清除旧的聚类结果标签
        for i in reversed(range(self.cluster_result_layout.count())): 
            widget = self.cluster_result_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()
        # 如果没有聚类结果，直接返回
        if not clustered_data:
            no_result = QLabel("暂无聚类结果")
            no_result.setFont(QFont("黑体", 8))
            no_result.setStyleSheet("color: gray;")
            self.cluster_result_layout.addWidget(no_result)
            return
        cluster_groups = defaultdict(list)
        for group_info, cluster_id in clustered_data:
            cluster_groups[cluster_id].append(group_info.Group)
        # 设置颜色映射（每个聚类序号对应不同颜色）
        color_map = {
            0: "#e6ffe6",  # 绿色
            1: "#e6f7ff",  # 蓝色
            2: "#fff3e0",  # 橙色
            3: "#f8bbd0",  # 粉色
            4: "#e1bee7",  # 紫色
            5: "#ffe0b2",  # 浅橙色
            6: "#cfd8dc",  # 灰蓝色
            7: "#fcbcbc",
            8: "#e0f9c5",
            9: "#fff9c4",
            10: "#c1f3fd",
        }
        # 添加聚类结果标签
        for cluster_id, groups in sorted(cluster_groups.items()):
            label_text = f"类别{cluster_id+1}:{'、'.join(map(str, sorted(groups)))}"
            label = QLabel(label_text)
            label.setFont(QFont("黑体", 9))
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            label.setFixedHeight(25)
            label.setStyleSheet(f"""
                background-color: {color_map.get(cluster_id, "#ffffff")};
                border: 1px solid #ddd;
                padding-left: 5px;
                border-radius: 3px;
            """)
            self.cluster_result_layout.addWidget(label)
    def select_files(self):
        """打开文件选择对话框"""
        file_dialog = QFileDialog()
        file_paths, _ = file_dialog.getOpenFileNames(
            self, 
            "选择Excel文件", 
            "", 
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if not file_paths:
            self.selected_files = []
            self.preview_data()
            QMessageBox.warning(self, "未选择文件", "未选择任何文件")
            return
        valid_files = []
        invalid_files = []
        for file_path in file_paths:
            try:
                df = pd.read_excel(file_path, nrows=0)  # 仅读取列名
                columns = df.columns.tolist()
                # 检查时间列是否存在（默认为第一列）
                time_column = columns[0] if len(columns) > 0 else None
                if time_column is None:
                    raise ValueError("时间列不存在（文件中没有列）")
                if not isinstance(time_column, str) or "时间" not in time_column and "Time" not in time_column:
                    raise ValueError(f"第一列 '{time_column}' 不是时间列（应包含 '时间' 或 'Time' 字样）")
                # 检查编码列是否存在（如 S1Eoding_1）
                individual_columns = [col for col in columns if col.startswith('S') and 'Eoding' in col]
                if not individual_columns:
                    raise ValueError("文件中没有找到编码列（如 S1Eoding_1）")
                # 文件检查通过
                valid_files.append(file_path)
            except Exception as e:
                # 文件检查失败
                invalid_files.append((file_path, str(e)))
        if valid_files:
            self.selected_files = valid_files
            self.preview_data()
            # 提取文件名
            invalid_file_names = '\n'.join([os.path.basename(f[0]) for f in invalid_files])
            message = (
                f"已选择 {len(valid_files)} 个符合要求的文件\n"
                f"以下文件不符合要求：\n{invalid_file_names}"
            )
            QMessageBox.information(self, "文件选择成功", message)
        else:
            self.selected_files = []
            self.result_table.setRowCount(0)
            if invalid_files:
                error_msg = "\n".join([f"{os.path.basename(f[0])}: {f[1]}" for f in invalid_files])
                QMessageBox.critical(
                    self,
                    "文件格式错误",
                    f"以下文件不符合要求：\n{error_msg}"
                )
            else:
                QMessageBox.warning(self, "未选择文件", "未选择任何符合要求的文件")
    def preview_data(self):
        all_data = []
        for file_idx, file_name in enumerate(self.selected_files):
            try:
                df = pd.read_excel(file_name)
                df['Group'] = file_idx + 1  # 使用文件顺序作为组别
                all_data.append(df)
            except Exception as e:
                print(f"读取文件失败: {file_name} - {str(e)}")
                continue
        if all_data:
            self.preview_df = pd.concat(all_data, ignore_index=True)
            self.display_preview(self.preview_df)
        else:
            self.result_table.setRowCount(0)
            QMessageBox.warning(self, "预览失败", "未成功读取任何有效数据")
    def display_preview(self, df, max_rows=5000):
        # 清除现有内容
        self.result_table.clear()
        # 限制显示行数，避免卡顿
        display_df = df.head(max_rows)
        # 确保 Group 列存在
        if 'Group' in display_df.columns:
            # 重新排序列：Group 放在最前面
            columns = ['Group'] + [col for col in display_df.columns if col != 'Group']
            display_df = display_df[columns]
        # 设置行和列
        self.result_table.setRowCount(display_df.shape[0])
        self.result_table.setColumnCount(display_df.shape[1])
        # 设置表头
        self.result_table.setHorizontalHeaderLabels([str(col) for col in display_df.columns])
        # 填充数据
        for i, row in display_df.iterrows():
            for j, value in enumerate(row):
                self.result_table.setItem(i, j, QTableWidgetItem(str(value)))
        # 自动调整列宽
        self.result_table.resizeColumnsToContents()
        # 提示信息
        QMessageBox.information(self, "数据预览", f"已预览前 {display_df.shape[0]} 行数据（完整数据将在分析后处理）")
    def run_analysis(self,encoding_config):
        if not hasattr(self, 'selected_files') or not self.selected_files:
            raise ValueError("请先选择要分析的Excel文件")
        try:
            all_data = []  # 创建一个空列表用于存储所有数据
            # 2. 处理数据
            # 读取所有Excel文件并合并数据
            student_maxnumber=self.student_maxcount
            student_minnumber=self.student_mincount
            for file_idx, file_name in enumerate(self.selected_files):
                try:
                    df = pd.read_excel(file_name)
                    for i in range(student_minnumber+1,student_maxnumber+1):
                        if f'S{i}Eoding_1' not in df.columns:  # 如果没有S3Eoding_1列
                            df[f'S{i}Eoding_1'] = None  # 添加S3Eoding_1列为None
                            df[f'S{i}Eoding_2'] = None  # 添加S3Eoding_2列为None
                            df[f'S{i}Eoding_3'] = None  # 添加S3Eoding_3列为None
                    # 添加组别列（基于文件顺序）
                    df['Group'] = file_idx+1  # 使用文件选择顺序作为组别
                    all_data.append(df)
                except Exception as e:
                    print(f"读取文件失败: {file_name} - {str(e)}")
                    continue
            combined_df = pd.concat(all_data, ignore_index=True)  # 将所有数据合并成一个DataFrame
            encoding_columns=self.get_encoding_columns()  
            dimension_count = len(encoding_config)
            # 提取组别
            groups = combined_df['Group'].unique()  # 获取所有唯一的组别
            group_stats_list = []  # 创建一个空列表用于存储每个小组的统计结果
            # 遍历每个小组
            for group in groups:  # 遍历每个唯一组别
                group_df = combined_df[combined_df['Group'] == group]  # 获取当前小组的数据
                counts = {col: {} for col in encoding_columns}  # 初始化计数字典
                for col in encoding_columns:  # 遍历每个编码列
                    col_counts = group_df[col].value_counts().to_dict()  # 计算每种编码类型的数量
                    for code, count in col_counts.items():  # 遍历每种编码类型及其数量
                        if code is not None:  # 如果编码类型不为空
                            if code not in counts[col]:  # 如果编码类型不在计数字典中
                                counts[col][code] = 0  # 初始化该编码类型的计数为0
                            counts[col][code] += count  # 累加该编码类型的计数
                row = {'Group': group}  # 创建一个字典存储当前小组的信息
                from collections import defaultdict
                total_counts = [0] * dimension_count  # 存储每个维度的总数
                # 计算各维度总数
                for dim_idx in range(dimension_count):
                    for stu_idx in range(student_maxnumber):
                        col = f'S{stu_idx+1}Eoding_{dim_idx+1}'
                        total_counts[dim_idx] += sum(counts.get(col, {}).values())
                for dim_idx, codes in enumerate(encoding_config):
                    for code in codes:
                        key = f'{code}_prop'
                        weighted_sum = 0
                    #计算每个学生的比例
                        for stu_idx in range(student_maxnumber):
                            col = f'S{stu_idx+1}Eoding_{dim_idx+1}'
                            code_count = counts.get(col, {}).get(code, 0)
                            weighted_sum += code_count / total_counts[dim_idx] if total_counts[dim_idx] else 0
                        row[key] = round(weighted_sum, 4)  # 保留4位小数
                group_stats_list.append(pd.DataFrame([row]))  # 将当前小组的统计结果添加到列表中
            # 合并所有统计结果
            group_stats = pd.concat(group_stats_list, ignore_index=True)  # 将所有统计结果合并成一个DataFrame
            # 提取特征列
            feature_columns = [col for col in group_stats.columns if '_prop' in col]  # 提取所有比例特征列
            X = group_stats[feature_columns].values  # 提取特征值
            # 标准化数据
            scaler = StandardScaler()  # 创建StandardScaler对象
            X_scaled = scaler.fit_transform(X)  # 标准化特征值
            n_clusters = self.n_clusters  # 设置聚类数目
            clusters = self.hierarchical_clustering(X_scaled, n_clusters=n_clusters, metric='euclidean', linkage='ward')  # 进行聚类
            # 绘制树状图
            self.plot_dendrogram(X_scaled, f'Hierarchical Clustering Dendrogram (n_clusters={n_clusters})')  # 绘制树状图
            # 显示结果
            clustered_data = list(zip(group_stats.itertuples(index=False), clusters))  # 将统计结果和聚类结果组合在一起
            self.display_results(clustered_data=clustered_data)
        except Exception as e:
            print(f"分析执行错误: {str(e)}")
            raise
    def canvas_edit(self):
        # 这里实现你的聚类分析逻辑
        # 获取编码配置
        encoding_config = []
        for dim in self.dimensions:
            codes = dim.get_codes()
            encoding_config.append(codes)
        # 执行数据处理和分析
        try:
            # 替换你的test1.py中的逻辑到这里
            self.canvas_show(encoding_config)
            # 显示结果
        except Exception as e:
            print("分析出错:", str(e))
    def canvas_show(self,encoding_config):
        if not hasattr(self, 'selected_files') or not self.selected_files:
            raise ValueError("请先选择要分析的Excel文件")
        try:
            all_data = []  # 创建一个空列表用于存储所有数据
            # 2. 处理数据
            # 读取所有Excel文件并合并数据
            student_maxnumber=self.student_maxcount
            student_minnumber=self.student_mincount
            for file_idx, file_name in enumerate(self.selected_files):
                try:
                    df = pd.read_excel(file_name)
                    for i in range(student_minnumber+1,student_maxnumber+1):
                        if f'S{i}Eoding_1' not in df.columns:  # 如果没有S3Eoding_1列
                            df[f'S{i}Eoding_1'] = None  # 添加S3Eoding_1列为None
                            df[f'S{i}Eoding_2'] = None  # 添加S3Eoding_2列为None
                            df[f'S{i}Eoding_3'] = None  # 添加S3Eoding_3列为None
                    # 添加组别列（基于文件顺序）
                    df['Group'] = file_idx+1  # 使用文件选择顺序作为组别
                    all_data.append(df)
                except Exception as e:
                    print(f"读取文件失败: {file_name} - {str(e)}")
                    continue
            combined_df = pd.concat(all_data, ignore_index=True)  # 将所有数据合并成一个DataFrame
            encoding_columns=self.get_encoding_columns()  
            dimension_count = len(encoding_config)
            # 提取组别
            groups = combined_df['Group'].unique()  # 获取所有唯一的组别
            group_stats_list = []  # 创建一个空列表用于存储每个小组的统计结果
            # 遍历每个小组
            for group in groups:  # 遍历每个唯一组别
                group_df = combined_df[combined_df['Group'] == group]  # 获取当前小组的数据
                counts = {col: {} for col in encoding_columns}  # 初始化计数字典
                for col in encoding_columns:  # 遍历每个编码列
                    col_counts = group_df[col].value_counts().to_dict()  # 计算每种编码类型的数量
                    for code, count in col_counts.items():  # 遍历每种编码类型及其数量
                        if code is not None:  # 如果编码类型不为空
                            if code not in counts[col]:  # 如果编码类型不在计数字典中
                                counts[col][code] = 0  # 初始化该编码类型的计数为0
                            counts[col][code] += count  # 累加该编码类型的计数
                row = {'Group': group}  # 创建一个字典存储当前小组的信息
                from collections import defaultdict
                total_counts = [0] * dimension_count  # 存储每个维度的总数
                # 计算各维度总数
                for dim_idx in range(dimension_count):
                    for stu_idx in range(student_maxnumber):
                        col = f'S{stu_idx+1}Eoding_{dim_idx+1}'
                        total_counts[dim_idx] += sum(counts.get(col, {}).values())
                for dim_idx, codes in enumerate(encoding_config):
                    for code in codes:
                        key = f'{code}_prop'
                        weighted_sum = 0
                    #计算每个学生的比例
                        for stu_idx in range(student_maxnumber):
                            col = f'S{stu_idx+1}Eoding_{dim_idx+1}'
                            code_count = counts.get(col, {}).get(code, 0)
                            weighted_sum += code_count / total_counts[dim_idx] if total_counts[dim_idx] else 0
                        row[key] = round(weighted_sum, 4)  # 保留4位小数
                group_stats_list.append(pd.DataFrame([row]))  # 将当前小组的统计结果添加到列表中
            # 合并所有统计结果
            group_stats = pd.concat(group_stats_list, ignore_index=True)  # 将所有统计结果合并成一个DataFrame
            # 提取特征列
            feature_columns = [col for col in group_stats.columns if '_prop' in col]  # 提取所有比例特征列
            X = group_stats[feature_columns].values  # 提取特征值
            # 标准化数据
            scaler = StandardScaler()  # 创建StandardScaler对象
            X_scaled = scaler.fit_transform(X)  # 标准化特征值
            n_clusters = self.n_clusters  # 设置聚类数目
            clusters = self.hierarchical_clustering(X_scaled, n_clusters=n_clusters, metric='euclidean', linkage='ward')  # 进行聚类
            # 绘制树状图
            self.plot_dendrogram_edit(X_scaled, f'Hierarchical Clustering Dendrogram (n_clusters={n_clusters})')
            # 显示结果
            clustered_data = list(zip(group_stats.itertuples(index=False), clusters))  # 将统计结果和聚类结果组合在一起
        except Exception as e:
            print(f"分析执行错误: {str(e)}")
    def hierarchical_clustering(self,X, n_clusters, metric='euclidean', linkage='ward'):
        clustering = AgglomerativeClustering(n_clusters=n_clusters, metric=metric, linkage=linkage)  # 创建AgglomerativeClustering对象
        clusters = clustering.fit_predict(X)  # 进行聚类并获取聚类结果
        return clusters  # 返回聚类结果
    # 绘制树状图
    def plot_dendrogram(self,X, title):
        self.figure.clear()
        # 创建子图
        ax = self.figure.add_subplot(111)
        group_count=self.group_count
        linked = linkage(X, method='ward')  # 计算链接矩阵
        dendrogram(linked,
               orientation='top',
               labels=[f'G{group}' for group in range(1,group_count+1)],
               distance_sort='descending',
               show_leaf_counts=True,
               ax=ax)  # 绘制树状图
        # 设置图表样式
        ax.set_title(title)
        ax.set_xlabel('Sample index')
        ax.set_ylabel('Distance')
        ax.grid(True)
        # 重绘 canvas
        self.canvas.draw()
    def plot_dendrogram_edit(self,X, title):
        group_count=self.group_count
        linked = linkage(X, method='ward')  # 计算链接矩阵
        plt.figure(figsize=(15, 10))  # 创建图形
        dendrogram(linked,
                orientation='top',
                labels=[f'G{group}' for group in range(1,group_count+1)],
                distance_sort='descending',
                show_leaf_counts=True)  # 绘制树状图
        plt.title(title)  # 设置标题
        plt.xlabel('Sample index')  # 设置x轴标签
        plt.ylabel('Distance')  # 设置y轴标签
        plt.show()  # 显示图形
    def copy_encoding_from_other_tab(self):
        """从编码标签页复制维度配置"""
        # 收集所有可用的编码标签页
        data_encoder_tabs = []
        for i in range(self.tabmodule.count()):
            widget = self.tabmodule.widget(i)
            if isinstance(widget, DataEncoderModule):
                tab_text = self.tabmodule.tabText(i)
                data_encoder_tabs.append((i, tab_text, widget))
        if not data_encoder_tabs:
            QMessageBox.warning(self, "无可用编码", "没有找到可复制的编码个体标签页")
            return
        # 弹出选择对话框
        items = [text for _, text, _ in data_encoder_tabs]
        item, ok = QInputDialog.getItem(self, "选择编码个体", 
                                    "请选择要复制的编码个体标签页：", items, 0, False)
        if not ok or not item:
            return
        selected_index = items.index(item)
        data_encoder = data_encoder_tabs[selected_index][2]
        # 执行维度复制
        self.copy_dimensions_from_data_encoder(data_encoder)
        QMessageBox.information(self, "复制成功", "维度和编码类型已复制")
    def copy_dimensions_from_data_encoder(self, data_encoder):
        """从指定的DataEncoderModule复制维度配置"""
        # 清除现有维度
        while self.dimensions:
            dim = self.dimensions.pop()
            dim.setParent(None)
            dim.deleteLater()
        # 复制每个维度
        for module in data_encoder.modules:
            self.add_dimension()  # 创建新维度
            new_dimension = self.dimensions[-1]
            new_dimension.dimension_number = len(self.dimensions)
            new_dimension.update_label()
            # 清除默认添加的空编码类型
            while new_dimension.codes:
                new_dimension.remove_code()
            # 复制编码类型
            for block_info in module.blocks:
                code_text = block_info['entry'].text().strip()
                if code_text:
                    new_dimension.add_code()
                    new_code_entry = new_dimension.codes[-1]['entry']
                    new_code_entry.setText(code_text)
class Dimension(QWidget):
    def __init__(self, parent=None, dimension_number=1):
        super().__init__(parent)
        self.parent = parent
        self.dimension_number = dimension_number
        self.codes = []  # 存储编码类型
        self.initUI()
    def initUI(self):
        # 整体布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        # 模块头部
        header_frame = QFrame()
        header_frame.setStyleSheet("background-color: #d0d9ff;")
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(5, 0, 5, 0)
        header_layout.setSpacing(0)
        self.label = QLabel()
        self.update_label()
        self.label.setStyleSheet("font-size: 12pt; font-family: 黑体;background-color: #d0d9ff;")
        self.label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        self.label.setMinimumWidth(100)
        # 添加可编辑的输入框
        self.edit_entry = QLineEdit()
        self.edit_entry.setStyleSheet("background-color: white;")
        self.edit_entry.setMaximumWidth(100)
        self.edit_entry.setFont(QFont("黑体", 12))
        self.edit_entry.setVisible(False)
        self.edit_entry.editingFinished.connect(self.save_label)
        self.add_code_btn = QToolButton()
        self.add_code_btn.setText('+')
        self.add_code_btn.setFixedSize(30, 30)
        self.add_code_btn.setStyleSheet("font-size: 12px; background-color: #d0d9ff;")
        self.add_code_btn.clicked.connect(self.add_code)
        self.remove_code_btn = QToolButton()
        self.remove_code_btn.setText('-')
        self.remove_code_btn.setFixedSize(30, 30)
        self.remove_code_btn.setStyleSheet("font-size: 12px; background-color: #d0d9ff;")
        self.remove_code_btn.clicked.connect(self.remove_code)
        header_layout.addWidget(self.label)
        header_layout.addWidget(self.edit_entry)
        header_layout.addStretch()
        header_layout.addWidget(self.add_code_btn)
        header_layout.addWidget(self.remove_code_btn)
        header_frame.setLayout(header_layout)
        # 编码类型容器
        self.code_container = QWidget()
        self.code_layout = QGridLayout()
        self.code_layout.setContentsMargins(5, 5, 5, 5)
        self.code_layout.setSpacing(5)
        self.code_container.setLayout(self.code_layout)
        # 事件处理
        header_frame.mouseDoubleClickEvent = self.edit_label
        # 添加到布局
        layout.addWidget(header_frame)
        layout.addWidget(self.code_container)
        # 初始添加一个编码类型
        self.add_code()
    def update_label(self):
        """更新维度标签"""
        self.label.setText(f'维度{self.dimension_number}')
    def edit_label(self, event):
        """双击编辑维度名称"""
        self.label.setVisible(False)
        self.edit_entry.setText(self.label.text())
        self.edit_entry.setVisible(True)
        self.edit_entry.setFocus()
    def save_label(self):
        """保存维度名称"""
        new_text = self.edit_entry.text()
        if new_text:
            self.label.setText(new_text)
            self.edit_entry.setVisible(False)
            self.label.setVisible(True)
    def add_code(self):
        """添加新编码类型"""
        index = len(self.codes)
        row = index // 2
        col = index % 2
        # 创建容器
        code_frame = QFrame()
        code_frame.setStyleSheet("background-color: #e7e9fd;")
        code_layout = QHBoxLayout()
        code_layout.setContentsMargins(0, 0, 0, 0)
        code_layout.setSpacing(5)
        # 输入框
        entry = QLineEdit()
        entry.setStyleSheet("background-color: white;")
        entry.setFont(QFont("Times New Roman", 11))
        entry.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        # 添加到布局
        code_layout.addWidget(entry)
        code_frame.setLayout(code_layout)
        # 布局管理
        self.code_layout.addWidget(code_frame, row, col)
        self.codes.append({'entry': entry, 'frame': code_frame})
    def remove_code(self):
        """删除最后一个编码类型"""
        if self.codes:
            code = self.codes.pop()
            code['frame'].deleteLater()
            self.rearrange_codes()
    def rearrange_codes(self):
        """重新排列剩余的编码类型"""
        for i, code in enumerate(self.codes):
            row = i // 2
            col = i % 2
            self.code_layout.addWidget(code['frame'], row, col)
    def get_codes(self):
        """获取所有编码类型"""
        return [code['entry'].text().strip() for code in self.codes if code['entry'].text().strip()]
class SequenceAnalysisModule(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
    def initUI(self):
        # 创建主窗口部件
        main_layout = QHBoxLayout(self)
        # 顶部操作区
        top_bar = QVBoxLayout()
        self.file_btn = QPushButton("选择文件")
        self.file_btn.setFont(QFont('黑体', 10))
        self.file_btn.setFixedHeight(30)
        self.file_btn.clicked.connect(self.select_file)
        top_bar.addWidget(self.file_btn)
        self.file_combo = QComboBox() # 文件历史下拉框
        self.file_combo.currentTextChanged.connect(self.plot_data)
        top_bar.addWidget(self.file_combo)
       # 修改后代码
        self.figure_label = QLabel()
        self.figure_label.setText("多通道序列图")
        self.figure_label.setFont(QFont('黑体', 12))
        self.figure = Figure(figsize=(12, 8))
        self.canvas = FigureCanvas(self.figure)
        # 创建图表容器并设置背景
        figure_container = QWidget()
        figure_container.setStyleSheet("background-color: #d0d9ff;")
        label_and_canvas_layout = QVBoxLayout(figure_container)
        label_and_canvas_layout.setContentsMargins(10, 10, 10, 10)
        label_and_canvas_layout.addWidget(self.figure_label)
        label_and_canvas_layout.addWidget(self.canvas)
        # 将容器加入布局
        top_bar.addWidget(figure_container)
        # Matplotlib画布
        content_area = QVBoxLayout()
        # 创建独立容器并设置背景
        label_btn_container = QWidget()
        label_btn_container.setStyleSheet("background-color: #a6baff;")
        label_btn_layout = QVBoxLayout(label_btn_container)
        label_btn_layout.setContentsMargins(10, 10, 10, 10)
        # 保留原有的 cacluate_btn 样式属性
        self.cacluate_btn = QPushButton("计算编码统计")
        #self.cacluate_btn.setStyleSheet(" background-color: #f8f8ff;")
        self.cacluate_btn.setFont(QFont('黑体', 10))
        self.cacluate_btn.setFixedHeight(30)
        self.cacluate_btn.setDisabled(True)
        self.cacluate_btn.clicked.connect(self.analyze)
        label_btn_layout.addWidget(self.cacluate_btn)
        # 保留原有的 caculate_label 样式属性
        self.caculate_label = QLabel("统计结果")
        self.caculate_label.setFont(QFont('黑体', 12))
        label_btn_layout.addWidget(self.caculate_label)
        # 将容器加入布局
        content_area.addWidget(label_btn_container)
        self.result_table = QTableWidget()
        self.result_table.setSortingEnabled(False)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.result_table.setStyleSheet("background-color: white;")
        content_area.addWidget(self.result_table)
        content_container = QWidget()
        content_container.setLayout(content_area)
        main_layout.addLayout(top_bar)
        main_layout.addWidget(content_container)  # 添加容器
        # 样式设置
        self.setStyleSheet("background-color: #f8f8ff;")
    def select_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择Excel文件", 
            "", 
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            # 检查是否已存在该路径
            index = self.file_combo.findText(file_path)
            if index >= 0:
                self.file_combo.setCurrentIndex(index)
            else:
                self.file_combo.addItem(file_path)
                self.file_combo.setCurrentIndex(self.file_combo.count() - 1)
            self.cacluate_btn.setEnabled(True)
        else:
            QMessageBox.critical(self, "未选择文件", f"请选择符合格式的文件")
            return
        try:
            df = pd.read_excel(file_path)
            columns = df.columns.tolist()
            # 检查时间列是否存在（默认为第一列）
            time_column = columns[0] if len(columns) > 0 else None
            if time_column is None:
                raise ValueError("时间列不存在（文件中没有列）")
            if not isinstance(time_column, str) or "Time" not in time_column:
                raise ValueError(f"第一列 '{time_column}' 不是时间列（应包含 'Time' 字样）")
            time_series = df[time_column]
            for idx, value in enumerate(time_series):
                # 处理NaN值
                if pd.isna(value):
                    continue
                # 标准化数据类型
                if isinstance(value, (int, float)):
                    # 数字类型直接通过
                    continue
                if isinstance(value, str):
                    value = value.strip()
                    # 匹配HH:MM:SS或MM:SS格式
                    if ':' in value:
                        parts = value.split(':')
                        if len(parts) in [2,3]:  # MM:SS 或 HH:MM:SS
                            raise ValueError(
                                f"时间格式错误（行{idx+2}：'{value}'）\n"
                                "时间列必须使用秒数表示\n"
                                "请将时间转换为总分或总秒数（例如：1分30秒→90秒）"
                            )
                else:
                    raise ValueError(
                        f"时间类型错误（行{idx+2}：类型{type(value).__name__}）\n"
                        "时间列必须为数值形式"
                    )
            # 检查编码列是否存在（如 S1Eoding_1）
            individual_columns = [col for col in columns if col.startswith('S') and 'Eoding_' in col]
            if not individual_columns:
                raise ValueError("文件中没有找到编码列（如 S1Eoding_1,列名需以S+学生序号+Eoding_+编码序号）")
        except Exception as e:
            QMessageBox.critical(self, "文件格式错误", f"文件不符合要求：\n{str(e)}")
            return
    def analyze(self, file_path):
            file_path = self.file_combo.currentText()
            df = pd.read_excel(file_path)
            # 提取S1/S2列
            # 提取列名
            columns = df.columns.tolist()
            time_column = columns[0]
            # 提取所有个体（S1/S2/S3...）
            individuals = set()
            for col in columns:
                if col.startswith('S') and col[1:][0].isdigit() and 'Eoding' in col:
                    individual = col.split('Eoding')[0]
                    individuals.add(individual)
            individuals = sorted(list(individuals))  # ['S1', 'S2', 'S3', ...]
            # 提取所有维度（Eoding_1/Eoding_2...）
            dims = set()
            for col in columns:
                if 'Eoding' in col:
                    dim = col.split('Eoding')[1]
                    if dim.startswith('_') and dim[1:].isdigit():
                        dims.add(dim)  # '_1', '_2', ...
            dims = sorted(list(dims), key=lambda x: int(x[1:]))  # ['_1', '_2', '_3', ...]
            # 调用分析函数
            results = self.analyze_encodings(df, individuals, dims)  # 生成分析结果
            # 更新结果显示
            self.display_results(results)
    def plot_data(self, file_path):
        try:
            df = pd.read_excel(file_path)
            columns = df.columns.tolist()
            time_column = columns[0]
            individuals = sorted(set(col.split('Eoding')[0] for col in columns if col.startswith('S') and 'Eoding' in col))
            dims = sorted(set(col.split('Eoding')[1] for col in columns if 'Eoding' in col and col.split('Eoding')[1].startswith('_')), key=lambda x: int(x[1:]))
            invalid_dims = []
            for dim in dims:
                all_empty = True
                for ind in individuals:
                    col_name = f"{ind}Eoding{dim}"
                    if col_name in df.columns:
                        # 检查列是否非空
                        if not df[col_name].isna().all() and not (df[col_name] == '').all():
                            all_empty = False
                            break
                if all_empty:
                    invalid_dims.append(dim)
            if invalid_dims:
                msg = "以下维度数据全空，请删除：\n" + "\n".join(invalid_dims)
                QMessageBox.warning(self, "空维度检测", msg)
            valid_dims = [dim for dim in dims if dim not in invalid_dims]
            if not valid_dims:
                raise ValueError("所有维度均为空，请检查数据完整性")
            dims = valid_dims
            self.figure.clear()
            axes = self.figure.subplots(len(dims), 1)
            if len(dims) == 1:
                axes = [axes]
            else:
                axes = axes.flatten()
            cmap_names = ['spring','summer','autumn','winter','cool','Wistia','hot','PuBuGn', 'YlOrBr', 'YlOrRd', 'Reds', 'Greens', 'Blues', 'Oranges', 'Purples', 'PuOr', 'BrBG', 'PRGn', 'PiYG', 'RdBu', 'RdGy', 'R']
            cmap_dict = {dim: cmap_names[i % len(cmap_names)] for i, dim in enumerate(dims)}
            for ax_idx, dim in enumerate(dims):
                ax = axes[ax_idx]
                all_events = []
                for ind in individuals:
                    col_name = f"{ind}Eoding{dim}"
                    if col_name in columns:
                        events = df[[time_column, col_name]].sort_values(by=time_column)
                        all_events.append((ind, col_name, events))
                if not all_events:
                    continue
                # 获取编码类型并构建颜色映射
                all_types = pd.concat([e[2][e[1]] for e in all_events]).unique()
                cmap = plt.get_cmap(cmap_dict[dim])
                color_map = {atype: cmap(i / (len(all_types) - 1)) for i, atype in enumerate(all_types)}
                # 均匀分布个体位置
                num_individuals = len(individuals)
                y_positions = {ind: (i + 0.5) / num_individuals for i, ind in enumerate(individuals)}
                # 优化 gap 计算
                gap = max(0.1 / num_individuals, 0.05)
                added_labels = set()
                for label, col_name, events in all_events:
                    prev_time = None
                    prev_type = None
                    for _, row in events.iterrows():
                        time = row[time_column]
                        type_val = row[col_name]
                        is_null = pd.isna(type_val) or type_val == ''
                        if prev_time is not None and prev_type is not None:
                            prev_is_null = (prev_type is None)
                            color = 'white' if prev_is_null else color_map.get(prev_type, 'black')
                            if label in y_positions:
                                y = y_positions[label]
                                ax.axvspan(
                                    prev_time, time,
                                    ymin=y - (0.5 - gap) / num_individuals,
                                    ymax=y + (0.5 - gap) / num_individuals,
                                    color=color,
                                    alpha=0.6
                                )
                                if not prev_is_null and prev_type not in added_labels:
                                    ax.axvspan(0, 0, ymin=0, ymax=0,
                                            color=color,
                                            alpha=0.6,
                                            label=prev_type)
                                    added_labels.add(prev_type)
                        prev_time = time
                        prev_type = None if is_null else type_val
                    if prev_time is not None:
                        curr_color = 'white' if prev_type is None else color_map.get(prev_type, 'black')
                        if label in y_positions:
                            y = y_positions[label]
                            ax.axvspan(
                                prev_time, df[time_column].max(),
                                ymin=y - (0.5 - gap) / num_individuals,
                                ymax=y + (0.5 - gap) / num_individuals,
                                color=curr_color,
                                alpha=0.6
                            )
                if y_positions:
                    for y in y_positions.values():
                        ax.axhspan(
                            y - (0.5 - gap) / num_individuals,
                            y + (0.5 - gap) / num_individuals,
                            xmin=0, xmax=1,
                            facecolor='white', edgecolor='none', zorder=-1
                        )
                # 设置子图样式
                ax.set_title(f"Eoding_{dim}", fontsize=10)
                ax.set_yticks([y_positions[ind] for ind in individuals])
                ax.set_yticklabels(individuals)
                ax.xaxis.set_major_locator(plt.MaxNLocator(nbins=10))
                ax.tick_params(axis='x', rotation=45)
                ax.grid(True, axis='x', linestyle='--', alpha=0.5)
            # 添加全局图例
            handles_all = []
            labels_all = []
            for ax in axes:
                handles, labels = ax.get_legend_handles_labels()
                handles_all.extend(handles)
                labels_all.extend(labels)
            by_label = dict(zip(labels_all, handles_all))
            plt.rcParams['font.sans-serif'] = ['SimHei']
            plt.rcParams['axes.unicode_minus'] = False
            self.figure.legend(
                by_label.values(),
                by_label.keys(),
                loc='upper right',
                bbox_to_anchor=(1, 1),
                fontsize='small',
                borderaxespad=0.
            )
            self.figure.tight_layout(rect=[0, 0, 0.85, 1])
            self.canvas.draw()
        except Exception as e:
            print(f"加载文件时出错: {str(e)}")
    def analyze_encodings(self, df: pd.DataFrame, individuals: list, dims: list):
        results = {
            'dimension_stats': {},
            'total_stats': {}
        }
        total_counts = defaultdict(int)
        for dim in dims:
            stats = {}
            for ind in individuals:
                # 根据实际列名调整构造方式
                col_name = f"{ind}Eoding{dim}"  # 若列名含下划线，如 S1Eoding_1
                # col_name = f"{ind}Eoding{dim[1:]}"  # 若列名不含下划线，如 S1Eoding1
                if col_name not in df.columns:
                    continue
                series = df[col_name]
                # 调试输出
                # 统计频次（可选 dropna=False）
                counts = series.value_counts(dropna=True)
                non_null_count = series.count()
                percent = (counts / non_null_count * 100).round(2)
                #print(f"统计结果: {counts}")
                stats[ind] = {
                    'counts': counts.to_dict(),
                    'percent': percent.to_dict()
                }
            # 合并统计
            combined_counts = {}
            total_samples_in_dim = 0
            for ind in individuals:
                col_name = f"{ind}Eoding{dim}"
                if col_name in df.columns:
                    total_samples_in_dim += df[col_name].count()
            for ind, data in stats.items():
                for code, count in data['counts'].items():
                    combined_counts[code] = combined_counts.get(code, 0) + count
            combined_percent = {
                code: round(count / total_samples_in_dim * 100, 2) if total_samples_in_dim > 0 else 0
                for code, count in combined_counts.items()
            }
            results['dimension_stats'][dim] = {
                'individual_stats': stats,
                'combined_counts': combined_counts,
                'combined_percent': combined_percent,
                'total_samples': total_samples_in_dim
            }
        return results
    def display_results(self, results):
        """将分析结果显示为表格形式，支持任意数量的个体和维度"""
        if not results.get('dimension_stats'):
            self.result_table.setRowCount(0)
            self.result_table.setColumnCount(0)
            self.result_table.setRowCount(1)
            self.result_table.setColumnCount(1)
            self.result_table.setItem(0, 0, QTableWidgetItem("暂无有效统计结果"))
            self.result_table.setHorizontalHeaderLabels(["信息"])
            return
        # 获取所有个体（假设所有维度的个体一致）
        individuals = []
        for dim in results['dimension_stats']:
            individuals = list(results['dimension_stats'][dim]['individual_stats'].keys())
            break  # 只需取一次
        # 构建 stat_items 和 stat_labels
        stat_items = []
        stat_labels = []
        for ind in individuals:
            stat_items.append((ind, 'counts'))
            stat_labels.append(f"{ind} 频次")
            stat_items.append((ind, 'percent'))
            stat_labels.append(f"{ind} 百分比(%)")
        stat_items.append(('combined', 'combined_counts'))
        stat_labels.append("合计频次")
        stat_items.append(('combined', 'combined_percent'))
        stat_labels.append("合计百分比(%)")
        # 构建表头：每个维度 + 编码组合
        headers = []
        for dim in sorted(results['dimension_stats']):
            codes = sorted(results['dimension_stats'][dim]['combined_counts'].keys())
            for code in codes:
                headers.append(f"{dim} - {code}")
        # 设置表格行列结构
        self.result_table.setRowCount(len(stat_labels))
        self.result_table.setColumnCount(len(headers))
        self.result_table.setVerticalHeaderLabels(stat_labels)
        self.result_table.setHorizontalHeaderLabels(headers)
        # 填充数据：按列（编码）填充
        col_index = 0
        for dim in sorted(results['dimension_stats']):
            stats = results['dimension_stats'][dim]
            codes = sorted(stats['combined_counts'].keys())
            for code in codes:
                for row_idx, (source, stat_type) in enumerate(stat_items):
                    try:
                        if source == 'combined':
                            if stat_type == 'combined_counts':
                                value = str(stats['combined_counts'].get(code, 0))
                            elif stat_type == 'combined_percent':
                                value = str(stats['combined_percent'].get(code, 0))
                            else:
                                value = "0"
                        else:
                            value = str(stats['individual_stats'][source][stat_type].get(code, 0))
                    except KeyError:
                        value = "0"
                    self.result_table.setItem(row_idx, col_index, QTableWidgetItem(value))
                col_index += 1
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 根据内容自动调整
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    window = CombinedApplication()
    # 获取当前脚本所在目录的绝对路径
    # 构建图片的绝对路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    logoimage_path = os.path.join(current_dir, 'logo-1种尺寸.png')
    window.setWindowIcon(QIcon(logoimage_path))  # 修改为相对路径
    window.show()
    sys.exit(app.exec_())