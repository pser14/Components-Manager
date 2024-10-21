import openpyxl
import pandas as pd
from PyQt6.QtWidgets import QMainWindow
from UI.creat import Ui_MainWindow
from qfluentwidgets import TeachingTip, InfoBarIcon


class Create_Ui(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setObjectName('a')
        self.wb = openpyxl.load_workbook('data/Components.xlsx')
        self.CompName = " "
        self.CompType = " "
        self.CompPackage = " "
        self.CompNum = 0
        self.items = self.wb.sheetnames
        self.comboBox_type.addItems(self.items)
        self.comboBox_type.setPlaceholderText("请选择器件类型")
        self.comboBox_type.setCurrentIndex(-1)
        self.Connect()

    def Callback(self):
        pass

    def Connect(self):
        self.pushButton_clear.clicked.connect(self.ClearText)
        self.pushButton_delete.clicked.connect(self.DeleteComp)
        self.pushButton_create.clicked.connect(self.WriteComp)
        self.comboBox_type.activated.connect(self.ComboBox)
        pass

    def GetText(self):
        self.CompName = self.lineEdit_name.text()
        self.CompPackage = self.lineEdit_package.text()
        self.CompNum = self.spinBox.value()
        self.CompType = self.comboBox_type.currentText()

    def WriteComp(self):
        self.GetText()
        if self.CompType != "" and self.CompName != "" and self.CompPackage != "":
            self.sheet = pd.read_excel('./data/Components.xlsx', sheet_name=self.CompType, engine='openpyxl')
            self.CompExist = self.sheet['器件名称']
            self.Comp_length = self.CompExist.dropna().shape[0]
            # print(f"总长度是：{self.Comp_length}")

            self.sheet.at[self.Comp_length, '器件名称'] = self.CompName
            self.sheet.at[self.Comp_length, '封装形式'] = self.CompPackage
            self.sheet.at[self.Comp_length, '数量'] = self.CompNum
            with pd.ExcelWriter('data/Components.xlsx', engine='openpyxl', mode='a',
                                if_sheet_exists='replace') as writer:
                self.sheet.to_excel(writer, sheet_name=self.CompType, index=False)
        else:
            self.ShowErrorCreate()

    def ClearText(self):
        self.lineEdit_name.clear()
        self.lineEdit_package.clear()
        pass

    def DeleteComp(self):
        self.GetText()
        if self.CompType != "" and self.CompName != "" and self.CompPackage != "":
            self.sheet = pd.read_excel('./data/Components.xlsx', sheet_name=self.CompType, engine='openpyxl')
            Rows_Drop = self.sheet[self.sheet['器件名称'] == self.CompName].index
            self.sheet = self.sheet.drop(Rows_Drop)
            with pd.ExcelWriter('data/Components.xlsx', engine='openpyxl', mode='a',
                                if_sheet_exists='replace') as writer:
                self.sheet.to_excel(writer, sheet_name=self.CompType, index=False)
        else:
            self.ShowErrorDelete()

    def ShowErrorCreate(self):
        TeachingTip.create(target=self.pushButton_create, icon=InfoBarIcon.ERROR, title="Error",
                           content="缺少所创建器件的属性",
                           isClosable=True, duration=2000, parent=self)

    def ShowErrorDelete(self):
        TeachingTip.create(target=self.pushButton_delete, icon=InfoBarIcon.ERROR, title="Error",
                           content="缺少所移除器件的属性",
                           isClosable=True, duration=2000, parent=self)

    def ComboBox(self):
        print("1")
        pass
