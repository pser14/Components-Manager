from PyQt6.QtWidgets import QMainWindow
from UI.search import Ui_MainWindow
from qfluentwidgets import InfoBarIcon, TeachingTip
import openpyxl
import pandas as pd


class Search_Ui(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.Connect()
        self.wb = openpyxl.load_workbook('data/Components.xlsx')
        self.CompPackage = " "
        self.CompNum = 0
        self.CompName = " "
        self.PickNum = 0
        self.CompType = " "
        self.items = self.wb.sheetnames
        self.comboBox.addItems(self.items)
        self.comboBox.setPlaceholderText("请选择器件类型")
        self.comboBox.setCurrentIndex(-1)

    def ClearText(self):
        self.lineEdit_name.clear()
        self.lineEdit_package.clear()
        pass

    def Connect(self):
        self.pushButton_delete.clicked.connect(self.ClearText)
        self.pushButton_search.clicked.connect(self.CompSearch)
        self.pushButton_pick.clicked.connect(self.CompPick)
        pass

    def CompPick(self):
        self.GetText()
        if self.CompType != "" and self.CompName != "" and self.CompPackage != "":
            self.sheet = pd.read_excel('./data/Components.xlsx', sheet_name=self.CompType, engine='openpyxl')
            print("1")
            column_values = {col: self.sheet[col].tolist() for col in self.sheet.columns}
            print(column_values)
            Comp_name_exist = column_values['器件名称']
            Comp_package_exist = column_values['封装形式']
            if self.CompName in Comp_name_exist and self.CompPackage in Comp_package_exist:
                TargeRow1 = set(self.sheet[(self.sheet['器件名称'] == self.CompName)].index.tolist())
                TargeRow2 = set(self.sheet[(self.sheet['封装形式'] == self.CompPackage)].index.tolist())
                TargeRowLast = tuple(TargeRow1 & TargeRow2)
                self.CompNum = self.sheet.loc[TargeRowLast, '数量']
                NumAfter = self.CompNum - self.PickNum
                if NumAfter < 0:
                    NumAfter = 0
                self.sheet.at[TargeRowLast, '数量'] = NumAfter
                self.lcdNumber.display(NumAfter)
                if NumAfter == 0:
                    self.ShowErrorZero()
                with pd.ExcelWriter('data/Components.xlsx', engine='openpyxl', mode='a',
                                    if_sheet_exists='replace') as writer:
                    self.sheet.to_excel(writer, sheet_name=self.CompType, index=False)
        else:
            self.ShowErrorPick()
        pass

    def CompSearch(self):
        self.GetText()
        if self.CompType != "" and self.CompName != "" and self.CompPackage != "":
            self.sheet = pd.read_excel('./data/Components.xlsx', sheet_name=self.CompType, engine='openpyxl')
            print("1")
            column_values = {col: self.sheet[col].tolist() for col in self.sheet.columns}
            print(column_values)
            Comp_name_exist = column_values['器件名称']
            Comp_package_exist = column_values['封装形式']
            print(Comp_name_exist)
            print(Comp_package_exist)
            if self.CompName in Comp_name_exist and self.CompPackage in Comp_package_exist:
                TargeRow1 = set(self.sheet[(self.sheet['器件名称'] == self.CompName)].index.tolist())
                TargeRow2 = set(self.sheet[(self.sheet['封装形式'] == self.CompPackage)].index.tolist())
                TargeRowLast = tuple(TargeRow1 & TargeRow2)
                self.CompNum = self.sheet.loc[TargeRowLast, '数量']
                self.lcdNumber.display(self.CompNum)
                pass
            else:
                self.NotExist()
        else:
            self.ShowErrorSearch()

    def GetText(self):
        self.CompName = self.lineEdit_name.text()
        self.CompPackage = self.lineEdit_package.text()
        self.PickNum = self.spinBox.value()
        self.CompType = self.comboBox.currentText()
        pass

    def ShowErrorSearch(self):
        TeachingTip.create(target=self.pushButton_search, icon=InfoBarIcon.ERROR, title="Error",
                           content="缺少所查找器件的属性",
                           isClosable=True, duration=2000, parent=self)

    def ShowErrorZero(self):
        self.GetText()
        TeachingTip.create(target=self.pushButton_pick, icon=InfoBarIcon.ERROR, title="Error",
                           content=f"{self.CompName}的数量不足",
                           isClosable=True, duration=2000, parent=self)

    def ShowErrorPick(self):
        self.GetText()
        TeachingTip.create(target=self.pushButton_pick, icon=InfoBarIcon.ERROR, title="Error",
                           content="缺少所拿取器件的属性",
                           isClosable=True, duration=2000, parent=self)

    def NotExist(self):
        TeachingTip.create(target=self.pushButton_search, icon=InfoBarIcon.ERROR, title="Error",
                           content="此元件尚未入库",
                           isClosable=True, duration=2000, parent=self)