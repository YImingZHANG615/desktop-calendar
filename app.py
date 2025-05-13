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

def get_data_file_path():
    """
    返回用于存储 events_data.json 的路径：
      - 如果程序打包后，返回 EXE 所在目录；
      - 否则返回源码所在目录。
    """
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, "events_data.json")

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
        self.load_events()

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
            except Exception as e:
                print("Failed to load events:", e)
                self.events = {}
        else:
            self.events = {}

    def save_events(self):
        """将事件信息保存到持久化路径下的 JSON 文件"""
        json_file = get_data_file_path()
        try:
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(self.events, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print("Failed to save events:", e)

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
        """右键弹出菜单：添加、删除、修改事件"""
        menu = QMenu()
        addAction = QAction("Add Event", self)
        removeAction = QAction("Remove Event", self)
        editAction = QAction("Edit Event Type", self)

        menu.addAction(addAction)
        menu.addAction(removeAction)
        menu.addAction(editAction)

        globalPos = self.calendar.mapToGlobal(pos)
        action = menu.exec_(globalPos)

        selected_date = self.calendar.selectedDate()
        date_str = selected_date.toString("yyyy-MM-dd")

        if action == addAction:
            event_desc, ok_desc = QInputDialog.getText(self, "Add Event", "Enter event description:")
            if not ok_desc or not event_desc:
                return
            event_type, ok_type = QInputDialog.getText(self, "Add Event", "Enter event type:")
            if not ok_type or not event_type:
                return
            color_name, ok_color = QInputDialog.getText(self, "Add Event", "Enter color (e.g. 'red' or '#FF0000'):")
            if not ok_color or not color_name:
                return
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
          - highlight=True 时：文字颜色使用自定义颜色（默认为 red），背景为淡黄色
          - highlight=False 时：恢复为黑字、白底
        """
        fmt = QTextCharFormat()
        if highlight:
            if not color:
                color = "red"
            fmt.setForeground(QColor(color))
            fmt.setBackground(QColor("yellow").lighter(160))
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
    
    def close_application(self):
        """真正退出应用程序"""
        try:
            # 保存事件数据
            self.save_events()
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
