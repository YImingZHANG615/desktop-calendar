# -*- coding: utf-8 -*-
import sys
import os
import json
import ctypes

from PyQt5.QtWidgets import (
    QApplication, QWidget, QCalendarWidget, QVBoxLayout,
    QInputDialog, QMenu, QAction, QToolTip, QSizePolicy,
    QSystemTrayIcon, QStyle, QMessageBox
)
from PyQt5.QtGui import (
    QPainter, QPainterPath, QColor, QRegion, QCursor, QFont,
    QTextCharFormat, QIcon
)
from PyQt5.QtCore import Qt, QDate, QPoint, QEvent, QLocale, QRectF

def resource_path(relative_path):
    """
    返回资源文件的绝对路径。
    如果程序打包后（sys.frozen 为 True），则文件位于 sys._MEIPASS 目录下，
    否则返回源码所在目录。
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_app_data_dir():
    """
    返回应用程序数据目录，使用用户的AppData目录
    """
    app_name = "桌面日历"
    if sys.platform.startswith('win'):
        # Windows: %APPDATA%\桌面日历
        app_data = os.path.join(os.environ['APPDATA'], app_name)
    elif sys.platform.startswith('darwin'):
        # macOS: ~/Library/Application Support/桌面日历
        app_data = os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', app_name)
    else:
        # Linux/Unix: ~/.桌面日历
        app_data = os.path.join(os.path.expanduser('~'), '.' + app_name)
    
    # 确保目录存在
    os.makedirs(app_data, exist_ok=True)
    return app_data

def get_data_file_path():
    """
    返回用于存储 events_data.json 的路径
    使用用户的AppData目录，确保具有写入权限
    """
    return os.path.join(get_app_data_dir(), "events_data.json")

def get_event_types_file_path():
    """
    返回用于存储 event_types.json 的路径
    使用用户的AppData目录，确保具有写入权限
    """
    return os.path.join(get_app_data_dir(), "event_types.json")

# ---------------------------
# 自定义日历控件
# ---------------------------
class MyCalendar(QCalendarWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setGridVisible(True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # 设置英文区域（保证月份、星期显示英文）
        self.setLocale(QLocale(QLocale.English))
        # 使用 Arial 字体
        self.setFont(QFont("Arial", 10))
        # 默认选中当前日期，并显示当前月份
        current = QDate.currentDate()
        self.setSelectedDate(current)
        self.setCurrentPage(current.year(), current.month())

    def wheelEvent(self, event):
        """利用鼠标滚轮切换月份"""
        delta = event.angleDelta().y()
        year = self.yearShown()
        month = self.monthShown()
        if delta > 0:
            month += 1
            if month > 12:
                month = 1
                year += 1
        else:
            month -= 1
            if month < 1:
                month = 12
                year -= 1
        self.setCurrentPage(year, month)
        event.accept()

# ---------------------------
# 主窗口：无边框、半透明、带圆角、支持拖动与缩放
# 并加载外部 ICO 文件作为程序图标
# ---------------------------
class CalendarWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("CalendarWidget")
        # 设置为工具窗口（不显示在任务栏）、无边框、始终置于底部
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnBottomHint)
        # 半透明背景，80% 不透明
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setWindowOpacity(0.8)
        self.setStyleSheet("""
            #CalendarWidget {
                border: 2px solid #aaa;
            }
            QCalendarWidget {
                background-color: transparent;
            }
        """)

        # 加载 ICO 图标文件（calendar.ico）并设置为应用及窗口图标
        icon_path = resource_path("calendar.ico")
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            QApplication.setWindowIcon(icon)
            self.setWindowIcon(icon)

        self.initUI()

        # 用于存储：日期字符串 -> { "type": ..., "description": ..., "color": ... }
        self.events = {}
        # 用于存储事件类型定义：类型名称 -> { "color": ... }
        self.event_types = {}
        self.load_events()
        self.load_event_types()

        # 拖动/缩放相关变量
        self.dragging = False
        self.resizing = False
        self.drag_position = QPoint()
        self.resize_margin = 15
        self._resize_start_rect = None
        self._resize_start_pos = None

        self.setMinimumSize(300, 300)
        self.refresh_highlight()
        
        # 创建系统托盘图标
        self.create_tray_icon()

    def initUI(self):
        # 创建自定义日历控件
        self.calendar = MyCalendar(self)
        self.calendar.customContextMenuRequested.connect(self.open_context_menu)
        
        # 设置窗口标题
        self.setWindowTitle("桌面日历")
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(self.calendar)
        self.setLayout(layout)

        # 点击日期时（若该日期有事件）显示提示信息
        self.calendar.clicked.connect(self.on_date_clicked)

        # 右键菜单：添加、删除、修改事件
        self.calendar.setContextMenuPolicy(Qt.CustomContextMenu)
        self.calendar.customContextMenuRequested.connect(self.open_context_menu)

    def load_events(self):
        """从持久化路径读取 JSON 文件中的事件信息"""
        json_file = get_data_file_path()
        if os.path.exists(json_file):
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    self.events = json.load(f)
                print(f"成功从 {json_file} 加载事件数据")
            except Exception as e:
                print(f"Failed to load events: {e}")
                self.events = {}
        else:
            print(f"事件数据文件不存在，将创建新文件: {json_file}")
            self.events = {}
            # 创建空文件
            self.save_events()

    def load_event_types(self):
        """加载事件类型定义"""
        json_file = get_event_types_file_path()
        if os.path.exists(json_file):
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    self.event_types = json.load(f)
                print(f"成功从 {json_file} 加载事件类型数据")
            except Exception as e:
                print(f"Failed to load event types: {e}")
                self.create_default_event_types()
        else:
            print(f"事件类型文件不存在，将创建默认类型: {json_file}")
            self.create_default_event_types()
    
    def create_default_event_types(self):
        """创建默认的事件类型"""
        # 默认创建几个基本类型
        self.event_types = {
            "会议": {"color": "#FF5733"},
            "生日": {"color": "#33FF57"},
            "假期": {"color": "#3357FF"},
            "纪念日": {"color": "#FF33A8"}
        }
        self.save_event_types()
        print("已创建默认事件类型")

    def save_events(self):
        """将事件信息保存到持久化路径下的 JSON 文件"""
        json_file = get_data_file_path()
        try:
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(self.events, f, ensure_ascii=False, indent=2)
            print(f"事件数据已保存到: {json_file}")
        except Exception as e:
            print(f"Failed to save events: {e}")
            print(f"尝试保存到路径: {json_file}")
            
            # 尝试创建目录
            try:
                os.makedirs(os.path.dirname(json_file), exist_ok=True)
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(self.events, f, ensure_ascii=False, indent=2)
                print(f"在创建目录后成功保存事件数据到: {json_file}")
            except Exception as e2:
                print(f"即使创建目录后仍然无法保存: {e2}")

    def save_event_types(self):
        """保存事件类型定义"""
        json_file = get_event_types_file_path()
        try:
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(self.event_types, f, ensure_ascii=False, indent=2)
            print(f"事件类型数据已保存到: {json_file}")
        except Exception as e:
            print(f"Failed to save event types: {e}")
            print(f"尝试保存到路径: {json_file}")
            
            # 尝试创建目录
            try:
                os.makedirs(os.path.dirname(json_file), exist_ok=True)
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(self.event_types, f, ensure_ascii=False, indent=2)
                print(f"在创建目录后成功保存事件类型数据到: {json_file}")
            except Exception as e2:
                print(f"即使创建目录后仍然无法保存事件类型: {e2}")

    def refresh_highlight(self):
        """
        为已添加事件的日期高亮。
        如果当前日期没有事件，则额外设置淡蓝背景和蓝色文字突出显示
        """
        # 先为有事件的日期设置高亮
        for date_str, info in self.events.items():
            date_obj = QDate.fromString(date_str, "yyyy-MM-dd")
            color = info.get("color", "red")
            self.highlight_date(date_obj, True, color=color)
        # 再为当前日期（如果未添加事件）设置高亮
        today = QDate.currentDate()
        today_str = today.toString("yyyy-MM-dd")
        if today_str not in self.events:
            fmt = QTextCharFormat()
            fmt.setForeground(QColor("blue"))
            fmt.setBackground(QColor("#ADD8E6"))
            self.calendar.setDateTextFormat(today, fmt)

    def on_date_clicked(self, date):
        """点击日期时，如果该日期有事件则显示提示信息"""
        date_str = date.toString("yyyy-MM-dd")
        if date_str in self.events:
            event_type = self.events[date_str].get("type", "")
            event_desc = self.events[date_str].get("description", "")
            QToolTip.showText(QCursor.pos(), f"Type: {event_type}\nDescription: {event_desc}", self)

    def open_context_menu(self, pos):
        """右键弹出菜单：添加、删除、修改事件，管理事件类型"""
        menu = QMenu()
        addAction = QAction("添加事件", self)
        removeAction = QAction("删除事件", self)
        editAction = QAction("编辑事件", self)
        
        # 添加事件类型管理子菜单
        typeMenu = QMenu("事件类型管理", self)
        addTypeAction = QAction("添加新类型", self)
        editTypeAction = QAction("编辑类型", self)
        deleteTypeAction = QAction("删除类型", self)
        typeMenu.addAction(addTypeAction)
        typeMenu.addAction(editTypeAction)
        typeMenu.addAction(deleteTypeAction)

        menu.addAction(addAction)
        menu.addAction(removeAction)
        menu.addAction(editAction)
        menu.addSeparator()
        menu.addMenu(typeMenu)

        globalPos = self.calendar.mapToGlobal(pos)
        action = menu.exec_(globalPos)

        selected_date = self.calendar.selectedDate()
        date_str = selected_date.toString("yyyy-MM-dd")
        
        # 处理事件类型管理
        if action == addTypeAction:
            self.add_event_type()
        elif action == editTypeAction:
            self.edit_event_type()
        elif action == deleteTypeAction:
            self.delete_event_type()

        if action == addAction:
            # 输入事件描述
            event_desc, ok_desc = QInputDialog.getText(self, "添加事件", "请输入事件描述:")
            if not ok_desc or not event_desc:
                return
                
            # 选择事件类型
            if self.event_types:
                # 如果有预定义的事件类型，则从列表中选择
                type_list = list(self.event_types.keys())
                type_list.append("自定义...")  # 添加自定义选项
                event_type, ok_type = QInputDialog.getItem(self, "选择事件类型", 
                                                     "请选择事件类型:", 
                                                     type_list, 0, False)
                if not ok_type or not event_type:
                    return
                    
                if event_type == "自定义...":
                    # 如果选择了自定义，则手动输入类型和颜色
                    event_type, ok_type = QInputDialog.getText(self, "添加事件", "请输入事件类型:")
                    if not ok_type or not event_type:
                        return
                    color_name, ok_color = QInputDialog.getText(self, "添加事件", 
                                                         "请输入颜色(如 'red' 或 '#FF0000'):")
                    if not ok_color or not color_name:
                        return
                else:
                    # 使用预定义类型的颜色
                    color_name = self.event_types[event_type]["color"]
            else:
                # 如果没有预定义类型，则手动输入
                event_type, ok_type = QInputDialog.getText(self, "添加事件", "请输入事件类型:")
                if not ok_type or not event_type:
                    return
                color_name, ok_color = QInputDialog.getText(self, "添加事件", 
                                                     "请输入颜色(如 'red' 或 '#FF0000'):")
                if not ok_color or not color_name:
                    return
            
            # 保存事件信息
            self.events[date_str] = {
                "type": event_type,
                "description": event_desc,
                "color": color_name
            }
            self.highlight_date(selected_date, True, color=color_name)
            self.save_events()
        elif action == removeAction:
            if date_str in self.events:
                del self.events[date_str]
                self.highlight_date(selected_date, False)
                self.save_events()
                # 如果删除的是当天事件，则重新设置当前日期的高亮
                if selected_date == QDate.currentDate():
                    fmt = QTextCharFormat()
                    fmt.setForeground(QColor("blue"))
                    fmt.setBackground(QColor("#ADD8E6"))
                    self.calendar.setDateTextFormat(selected_date, fmt)
        elif action == editAction:
            if date_str in self.events:
                current_type = self.events[date_str].get("type", "")
                new_type, ok = QInputDialog.getText(self, "Edit Event Type", "Enter new event type:", text=current_type)
                if ok and new_type:
                    self.events[date_str]["type"] = new_type
                    self.save_events()

    def highlight_date(self, date, highlight, color=None):
        """
        设置指定日期的格式：
          - highlight=True 时：整个日期格子高亮，使用自定义颜色作为背景色，文字为白色
          - highlight=False 时：恢复为黑字、白底
        """
        fmt = QTextCharFormat()
        if highlight:
            if not color:
                color = "red"
            # 将文字设置为白色，增强可读性
            fmt.setForeground(QColor("white"))
            # 设置背景色为指定颜色，稍微调淡以增强视觉效果
            bg_color = QColor(color)
            # 确保颜色不会太深导致文字不清晰
            if bg_color.lightness() > 200:  # 如果颜色太浅
                fmt.setForeground(QColor("black"))  # 使用黑色文字
            # 设置背景色填充整个格子
            fmt.setBackground(bg_color)
        else:
            fmt.setForeground(Qt.black)
            fmt.setBackground(Qt.white)
        self.calendar.setDateTextFormat(date, fmt)

    # --- 自绘部分，实现圆角背景 ---
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        rect = QRectF(self.rect())
        path = QPainterPath()
        path.addRoundedRect(rect, 30, 30)
        painter.fillPath(path, QColor(255, 255, 255, 240))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), 30, 30)
        region = QRegion(path.toFillPolygon().toPolygon())
        self.setMask(region)

    # --- 重写鼠标事件，实现拖动与右下角缩放 ---
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            pos = event.pos()
            rect = self.rect()
            if pos.x() >= rect.width() - self.resize_margin and pos.y() >= rect.height() - self.resize_margin:
                self.resizing = True
                self._resize_start_rect = self.geometry()
                self._resize_start_pos = event.globalPos()
            else:
                self.dragging = True
                self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.resizing:
            delta = event.globalPos() - self._resize_start_pos
            new_width = self._resize_start_rect.width() + delta.x()
            new_height = self._resize_start_rect.height() + delta.y()
            if new_width < self.minimumWidth():
                new_width = self.minimumWidth()
            if new_height < self.minimumHeight():
                new_height = self.minimumHeight()
            self.resize(new_width, new_height)
            event.accept()
        elif self.dragging:
            self.move(event.globalPos() - self.drag_position)
            event.accept()
        else:
            rect = self.rect()
            if event.pos().x() >= rect.width() - self.resize_margin and event.pos().y() >= rect.height() - self.resize_margin:
                self.setCursor(Qt.SizeFDiagCursor)
            else:
                self.setCursor(Qt.ArrowCursor)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.dragging = False
        self.resizing = False
        self.setCursor(Qt.ArrowCursor)
        event.accept()

    def eventFilter(self, source, event):
        if source == self.calendar:
            if event.type() == QEvent.MouseButtonPress:
                pos = self.mapFromGlobal(event.globalPos())
                rect = self.rect()
                if pos.x() >= rect.width() - self.resize_margin and pos.y() >= rect.height() - self.resize_margin:
                    self.resizing = True
                    self._resize_start_rect = self.geometry()
                    self._resize_start_pos = event.globalPos()
                    return True
                else:
                    if pos.y() < 20:
                        self.dragging = True
                        self.drag_position = event.globalPos() - self.frameGeometry().topLeft()
                        return True
            elif event.type() == QEvent.MouseMove:
                if self.resizing:
                    delta = event.globalPos() - self._resize_start_pos
                    new_width = self._resize_start_rect.width() + delta.x()
                    new_height = self._resize_start_rect.height() + delta.y()
                    if new_width < self.minimumWidth():
                        new_width = self.minimumWidth()
                    if new_height < self.minimumHeight():
                        new_height = self.minimumHeight()
                    self.resize(new_width, new_height)
                    return True
                elif self.dragging:
                    self.move(event.globalPos() - self.drag_position)
                    return True
            elif event.type() == QEvent.MouseButtonRelease:
                self.resizing = False
                self.dragging = False
                return True
        return super().eventFilter(source, event)
        
    def create_tray_icon(self):
        """创建系统托盘图标及菜单"""
        # 创建托盘图标
        self.tray_icon = QSystemTrayIcon(self)
        
        # 使用程序图标作为托盘图标
        icon_path = resource_path("calendar.ico")
        if os.path.exists(icon_path):
            self.tray_icon.setIcon(QIcon(icon_path))
        else:
            # 如果找不到图标文件，使用系统自带的日历图标
            self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        
        # 创建托盘菜单
        self.tray_menu = QMenu()
        
        # 添加菜单项
        self.show_action = QAction("显示日历", self)
        self.hide_action = QAction("隐藏日历", self)
        self.exit_action = QAction("退出", self)
        
        # 连接信号
        self.show_action.triggered.connect(self.show)
        self.hide_action.triggered.connect(self.hide)
        self.exit_action.triggered.connect(self.close_application)
        
        # 添加到菜单
        self.tray_menu.addAction(self.show_action)
        self.tray_menu.addAction(self.hide_action)
        self.tray_menu.addSeparator()
        self.tray_menu.addAction(self.exit_action)
        
        # 设置托盘菜单
        self.tray_icon.setContextMenu(self.tray_menu)
        
        # 托盘图标双击事件
        self.tray_icon.activated.connect(self.tray_icon_activated)
        
        # 显示托盘图标
        self.tray_icon.show()
        
        # 设置托盘提示信息
        self.tray_icon.setToolTip("桌面日历")
    
    def tray_icon_activated(self, reason):
        """托盘图标被激活时的处理"""
        if reason == QSystemTrayIcon.DoubleClick:
            # 双击托盘图标时，如果窗口隐藏则显示，如果窗口显示则隐藏
            if self.isVisible():
                self.hide()
            else:
                self.show()
    
    def closeEvent(self, event):
        """重写关闭事件，点击关闭按钮时最小化到托盘而不是退出"""
        try:
            # 隐藏窗口
            self.hide()
            # 显示托盘消息提示
            self.tray_icon.showMessage(
                "桌面日历",
                "应用已最小化到系统托盘，双击托盘图标可以再次显示窗口。",
                QSystemTrayIcon.Information,
                2000
            )
            # 忽略关闭事件，不真正关闭窗口
            event.ignore()
        except Exception as e:
            print(f"关闭事件处理出错: {e}")
            event.ignore()
    
    def add_event_type(self):
        """添加新的事件类型"""
        type_name, ok_name = QInputDialog.getText(self, "添加事件类型", "请输入事件类型名称:")
        if not ok_name or not type_name:
            return
            
        if type_name in self.event_types:
            QMessageBox.warning(self, "类型已存在", f"事件类型 '{type_name}' 已经存在!")
            return
            
        color_name, ok_color = QInputDialog.getText(self, "添加事件类型", 
                                              "请输入颜色(如 'red' 或 '#FF0000'):", 
                                              text="#3366FF")
        if not ok_color or not color_name:
            return
            
        # 添加新类型
        self.event_types[type_name] = {"color": color_name}
        self.save_event_types()
        QMessageBox.information(self, "添加成功", f"事件类型 '{type_name}' 已成功添加!")
    
    def edit_event_type(self):
        """编辑现有事件类型"""
        if not self.event_types:
            QMessageBox.information(self, "无类型", "当前没有事件类型可编辑")
            return
            
        # 选择要编辑的类型
        type_name, ok = QInputDialog.getItem(self, "选择类型", "选择要编辑的事件类型:", 
                                         list(self.event_types.keys()), 0, False)
        if not ok or not type_name:
            return
            
        # 编辑颜色
        current_color = self.event_types[type_name].get("color", "#FF0000")
        color_name, ok_color = QInputDialog.getText(self, "编辑颜色", 
                                              "请输入新的颜色(如 'red' 或 '#FF0000'):", 
                                              text=current_color)
        if not ok_color or not color_name:
            return
            
        # 更新类型
        self.event_types[type_name]["color"] = color_name
        self.save_event_types()
        
        # 更新使用该类型的所有事件
        for date_str, event in self.events.items():
            if event.get("type") == type_name:
                event["color"] = color_name
        
        self.save_events()
        self.refresh_highlight()
        QMessageBox.information(self, "更新成功", f"事件类型 '{type_name}' 已成功更新!")
    
    def delete_event_type(self):
        """删除事件类型"""
        if not self.event_types:
            QMessageBox.information(self, "无类型", "当前没有事件类型可删除")
            return
            
        # 选择要删除的类型
        type_name, ok = QInputDialog.getItem(self, "选择类型", "选择要删除的事件类型:", 
                                         list(self.event_types.keys()), 0, False)
        if not ok or not type_name:
            return
            
        # 确认删除
        reply = QMessageBox.question(self, "确认删除", 
                                 f"确定要删除事件类型 '{type_name}' 吗?", 
                                 QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # 删除类型
            del self.event_types[type_name]
            self.save_event_types()
            QMessageBox.information(self, "删除成功", f"事件类型 '{type_name}' 已成功删除!")
    
    def close_application(self):
        """真正退出应用程序"""
        try:
            # 保存事件数据
            self.save_events()
            # 保存事件类型数据
            self.save_event_types()
            # 隐藏托盘图标
            self.tray_icon.hide()
            # 退出应用
            QApplication.quit()
        except Exception as e:
            print(f"退出应用时出错: {e}")
            # 强制退出
            sys.exit(0)

def add_to_startup():
    """在 Windows 下写入注册表实现开机自启"""
    if sys.platform.startswith("win"):
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Run",
                                 0, winreg.KEY_ALL_ACCESS)
            
            # 获取当前程序路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的EXE
                app_path = sys.executable
            else:
                # 如果是源码运行
                app_path = os.path.abspath(sys.argv[0])
                
            # 写入注册表
            winreg.SetValueEx(key, "桌面日历", 0, winreg.REG_SZ, app_path)
            winreg.CloseKey(key)
            print("已添加到开机自启动")
        except Exception as e:
            print(f"添加到开机自启失败: {e}")
            # 尝试使用备用方法
            try:
                startup_folder = os.path.join(os.environ["APPDATA"], 
                                          r"Microsoft\Windows\Start Menu\Programs\Startup")
                if getattr(sys, 'frozen', False):
                    # 创建快捷方式
                    import win32com.client
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut = shell.CreateShortCut(os.path.join(startup_folder, "桌面日历.lnk"))
                    shortcut.Targetpath = sys.executable
                    shortcut.WorkingDirectory = os.path.dirname(sys.executable)
                    shortcut.save()
                    print("已创建开机自启快捷方式")
            except Exception as e2:
                print(f"备用自启动方法也失败: {e2}")

if __name__ == "__main__":
    # 默认启用开机自启动
    add_to_startup()
    app = QApplication(sys.argv)
    
    # 检查系统是否支持系统托盘
    if not QSystemTrayIcon.isSystemTrayAvailable():
        QMessageBox.critical(None, "系统托盘不可用",
                           "在此系统上无法使用系统托盘功能。")
        sys.exit(1)
    
    # 设置退出行为：当最后一个窗口关闭时，不自动退出应用
    QApplication.setQuitOnLastWindowClosed(False)
    
    win = CalendarWidget()
    win.resize(400, 400)
    win.show()
    sys.exit(app.exec_())
