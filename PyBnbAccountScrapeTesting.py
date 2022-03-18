import os
import shutil
import tabula
import pandas as pd
import openpyxl
from openpyxl import Workbook
import csv
import undetected_chromedriver
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

ProjectPath = os.path.dirname(os.path.realpath(__file__))
TemplateExcelPath  = ProjectPath + r'\tempvlook_full.xlsx'
Pdf = r'2020-23700076.pdf'
i = 1
CompanyFolderPath = r'C:\Users\flyin\Documents\03. Projects\06. Autommation\4. Python - fy\PyBnbAccountScrape\Output\LIEDEKERKE WOLTERS WAELBROECK KIRKPATRICK'
wb = openpyxl.load_workbook(TemplateExcelPath)

def remove_empty(line):
    result = []
    for i in range(len(line)):
        if line[i] != "":
            result.append(line[i])
    return result

def ExtractPDF(Pdf,i, CompanyFolderPath, destinationwb):
    PDFPath = CompanyFolderPath+'\\'+str(Pdf)
    OutputCSVPath = CompanyFolderPath+'\\scrapeCSV'+str(i)+'.csv'

    tabula.convert_into(PDFPath, OutputCSVPath, pages='all', output_format="csv", stream=True)

    ws = destinationwb['Acc('+str(i)+')']
    with open(OutputCSVPath) as CSV:
        reader = csv.reader(CSV, delimiter=',')
        for row in reader:
            row = remove_empty(row)
            for i in range(len(row)):
                row[i] = row[i].replace(".","")
                try:
                    row[i] = float(row[i])
                except:
                    pass
            ws.append(row)


ExtractPDF(Pdf, i, CompanyFolderPath, wb)
