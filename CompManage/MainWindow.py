import sys
from creat_ui import Create_Ui
from search_ui import Search_Ui
from qfluentwidgets import NavigationItemPosition, FluentWindow, SubtitleLabel, setFont, SplitFluentWindow
from qfluentwidgets import FluentIcon as FIF
from PyQt6.QtGui import QIcon, QColor
from PyQt6.QtWidgets import QApplication, QWidget


class Demo(SplitFluentWindow):
    def __init__(self):
        super().__init__()
        self.resize(700, 200)
        self.setWindowIcon(QIcon('CompManage/Icon/add-circle.png'))
        self.setWindowTitle('CompManage')
        self.titleBar.maxBtn.hide()
        self.titleBar.setDoubleClickEnabled(False)
        self.setCustomBackgroundColor(QColor(242, 242, 242), QColor(25, 33, 42))
        self.CreateWindow = Create_Ui()
        self.SearchWindow = Search_Ui()

        self.addSubInterface(self.SearchWindow, FIF.SEARCH, '查找器件')
        self.addSubInterface(self.CreateWindow, FIF.ADD, '新建器件')


app = QApplication(sys.argv)
w = Demo()
w.show()
app.exec()
