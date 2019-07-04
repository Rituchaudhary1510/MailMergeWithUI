# -*- coding: utf-8 -*-
"""
Created on Thu Jul  4 22:31:38 2019

@author: YogiRitu
"""


from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from mailmerge import *
import xlrd
import os

import urllib.request
import sys



class Merger(QDialog):
    def __init__(self):
        QDialog.__init__(self)

        layout= QGridLayout()
        self.label1= QLabel("Enter Doc File:")
        self.file1=QLineEdit()
        btn1=QPushButton("Browse")
        self.label2=QLabel("Enter Excel File:")
        self.file2=QLineEdit()
        btn2=QPushButton("Browse")
        self.label3=QLabel("Enter Output File:")
        self.file3= QLineEdit()
        btn3= QPushButton("Save As")
        button=QPushButton("Merge")

        self.file1.setPlaceholderText("Docx file path")
        self.file2.setPlaceholderText("Excel file path")
        self.file3.setPlaceholderText("Output file path")
        # self.progress.setAlignment(Qt.AlignHCenter)

        layout.addWidget(self.label1)
        layout.addWidget(self.file1,0,1)
        layout.addWidget(btn1,0,2)
        layout.addWidget(self.label2)
        layout.addWidget(self.file2)
        layout.addWidget(btn2,1,2)
        layout.addWidget(self.label3)
        layout.addWidget(self.file3)
        layout.addWidget(btn3,2,2)
        layout.addWidget(button,3,1)

        self.setLayout(layout)
        self.setWindowTitle("FileMerger")
        btn1.clicked.connect(self.openFileNameDialog)
        btn2.clicked.connect(self.openFileNamesDialog)
        btn3.clicked.connect(self.saveFileDialog)
        button.clicked.connect(self.merge)
        
    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Doc Files (*.docx)", options=options)
        if fileName:
            print(fileName)
            self.file1.setText(fileName)
            return fileName
            
    def openFileNamesDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Excel Files (*.xlsx)", options=options)
        if fileName:
            print(fileName)
            self.file2.setText(fileName)
            return fileName
             
    def saveFileDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","Doc Files (*.docx)", options=options)
        if fileName:
            print(fileName)
            self.file3.setText(fileName)
    
    def merge(self):
        current_row = 0
        sheet_num = 0
        print("Reading .docx file.")
        file1=self.file1.text()
        document =MailMerge(file1)
        print(document.get_merge_fields())

            # path to the file you want to extract data from
            # src = sys.argv[2]
        file2=self.file2.text()
        print("Reading .xlsx file.")
        book = xlrd.open_workbook(file2)
            # book = xlrd.open_workbook(src)
        if ((book.nsheets >= sheet_num + 1) == False):
            print("Unable to find sheet number provided.")
            sys.exit()
            # select the sheet that the data resids in
        work_sheet = book.sheet_by_index(sheet_num)
        finalList = []
        headers = []
            # get the total number of rows
        num_rows = work_sheet.nrows

            # Format required for mail merge is :
            # List [
            #  {Dictonary},
            #  {Dictonary},
            #  ..
            # ]
        print("Preparing to merge.")
        while current_row < num_rows:
            dictVal = dict()
            if (current_row == 0):
                for col in range(work_sheet.ncols):
                    headers.append(work_sheet.cell_value(current_row, col))
            else:
                for col in range(work_sheet.ncols):
                    dictVal.update({headers[col]: work_sheet.cell_value(current_row, col)})
            if (current_row != 0):
                finalList.append(dictVal)
            current_row += 1
        print(finalList)
        print("Merge operation started.")
        document.merge_pages(finalList)
        print("Saving ouput file.")
        file3=self.file3.text()
        document.write(file3)

        print("Operation completed successfully.")
        self.closeIt()
        self.messagebox()
    
    def messagebox(self):
        msg= QMessageBox.about(self, "Information", "Merge completed successfully")
        
        
    def closeIt(self): 
        self.close()


app= QApplication([])
windows=Merger()
windows.show()
app.exec_()




















