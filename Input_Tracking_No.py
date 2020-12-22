#import datetime, time, os, re, win32com.client, shutil, telepot
import datetime, win32com.client
#import pyautogui
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

###################################################################
#   Define List, Variable
###################################################################
#form_class = uic.loadUiType("D:\\03_Study\\01_Python\\01_Code\\02_Auto\\Trackingnumber_Convertor.ui")[0]
form_class = uic.loadUiType("Input_Tracking_No.ui")[0]
app = 0

###################################################################
#   각 열의 위치를 지정
#   문자에 대한 ASCII 코드 변환 후 64를 빼서 A부터 1로 카운트
###################################################################
SOURCE_NAME = ord('A') - 64 # 수취인명(A열)
SOURCE_TEL_NO = ord('B') - 64 # 수취인 연락처1(B열)
SOURCE_TRACKING_NO = ord('L') - 64 # 한진택배 송장번호(L열)

TARGET_NAME = ord('K') - 64 # 수취인명
TARGET_TEL_NO = ord('O') - 64 + 26 # 수취인 연락처1(AO)
TARGET_TRACKING_NO = ord('F') - 64 # 스마트스토어 송장번호 송장번호

TARGET_FORWARDING_NAME = ord('E') - 64 # 스마트스토어 택배사 이름 : 모두 한진택배로 기입 필요

###################################################################
#   Insection File Name Only
###################################################################
def Insection_Filename(full_path):
    x = full_path.split('/')
    x.reverse()
    ret_filename = x[0]
    return ret_filename

###################################################################
#   Working Directory
###################################################################
def Working_code(source, target):
    ### Excel File 정보 ###
    Source_Excel_PATH = source    
    Target_Excel_PATH = target
    
    excel = win32com.client.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(Source_Excel_PATH)
    ws = wb.Worksheets('원본') # 추후 Exception Code 추가 필요. To Do !!!
    
    wb1 = excel.Workbooks.Open(Target_Excel_PATH)
    ws1 = wb1.Worksheets('발주발송관리')

    ### Looking for number of item from Excel file for Source ###
    row = 1

    while True:
        cell_value = ws.Cells(row, 1).Value

        if cell_value == None:
            source_last_row = row - 1
            break

        row += 1

    ### Looking for number of item from Excel file for Target ###
    row = 3

    while True:
        cell_value = ws1.Cells(row, 1).Value

        if cell_value == None:
            item_number = row - 3
            target_last_row = row - 1
            break

        row += 1   

    #############################################
    ### Data copy from Source to Target #########
    #############################################

    for i in range(source_last_row):
        for j in range(item_number):
            if ((ws1.Cells(j+3, TARGET_NAME).Value) == (ws.Cells(i+1, SOURCE_NAME).Value)) and ((ws1.Cells(j+3, TARGET_TEL_NO).Value) == (ws.Cells(i+1, SOURCE_TEL_NO).Value)): # 수취인명 : Target 데이터는 3행 부터 시작 / Source 데이터는 1행 부터 시작
                ws1.Cells(j+3, TARGET_FORWARDING_NAME).Value = '한진택배' 
                ws1.Cells(j+3, TARGET_TRACKING_NO).Value = ws.Cells(i+1, SOURCE_TRACKING_NO).Value
                
    ### 엑셀 파일을 저장 후 종료
    wb1.Save()
    wb1.Close()
    wb.Close()
    #excel.Quit()


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        self.Source_File_Path = 0
        self.Target_File_Path = 0
        self.Source_input = False
        self.Target_input = False
        
        self.pushButton.clicked.connect(self.pushButtonClicked)
        self.pushButton_2.clicked.connect(self.pushButtonClicked2)
        self.pushButton_3.clicked.connect(self.pushButtonClicked3)
        self.pushButton_4.clicked.connect(app.quit)

    def pushButtonClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        self.Source_File_Path = fname[0] # 파일 이름을 포함한 전체 경로
        Source_Filename = Insection_Filename(self.Source_File_Path) # 파일 이름
        
        if 'xlsx' in Source_Filename:
            self.textEdit.setText(Source_Filename)
            self.label.setText('한진택배 파일 선택 완료')
            self.Source_input = True
        else:
            #self.textEdit.clearText()
            self.label.setText('한진택배 파일을 선택해주세요 !!!')
            self.Source_input = False

    def pushButtonClicked2(self):
        fname = QFileDialog.getOpenFileName(self)
        self.Target_File_Path = fname[0] # 파일 이름을 포함한 전체 경로
        Target_Filename = Insection_Filename(self.Target_File_Path) # 파일 이름
        
        if '스마트스토어' in Target_Filename:
            self.textEdit_2.setText(Target_Filename)
            self.label.setText('스마트스토어 파일 선택 완료')
            self.Target_input = True
        else:
            #self.textEdit_2.clearText()
            self.label.setText('스마트스토어 파일을 선택해주세요 !!!')
            self.Target_input = False

    def pushButtonClicked3(self):
        if (self.Source_input == True) and (self.Target_input == True):            
            self.label.setText('파일 변환 진행 중 ...')
            Working_code(self.Source_File_Path, self.Target_File_Path)
            self.label.setText('완료 ! 파일은 한진택배_Template 파일 폴더에 저장됨 !! 종료 버튼을 누르세요 !!')
            
            self.Source_input = False
            self.Target_input = False
        else:
            if self.Source_input == True:
                self.label.setText('한진택배 파일을 선택해주세요 !!')
            elif self.Target_input == True:                
                self.label.setText('스마트스토어 파일을 선택해주세요 !!')
            else:
                self.label.setText('한진택배 파일과 스마트스토어 파일을 선택해주세요 !!')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()