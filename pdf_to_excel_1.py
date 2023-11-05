import openpyxl
import glob
import tabula # for pdf to excel
import aspose.pdf as ap
from pyexcel.cookbook import merge_all_to_a_book
# import pyexcel.ext.xlsx # no longer required if you use pyexcel >= 0.2.2 

from openpyxl import Workbook
import csv

pdf_name="Invoice160.pdf"
xlsxname="agile_try.xlsx"
csvname="agile_try.csv"
input_agile_path = "input/pdf/"
output_excel_path = "output/excel/"


def pdf_to_excel(pdf_name,xlsxname):
    # Read PDF File
    df = tabula.read_pdf(input_agile_path+pdf_name, pages = 1)[0]

    # Convert into Excel File
    df.to_excel(output_excel_path+xlsxname)

#pdf_to_excel(pdf_name,xlsxname)

def pdf_to_excel_2(pdf_name,xlsxname):
    # Open PDF document
    document = ap.Document(input_agile_path+pdf_name)
    save_option = ap.ExcelSaveOptions()
    # Save the file into MS Excel format
    document.save(output_excel_path+xlsxname, save_option)

#pdf_to_excel_2(pdf_name,xlsxname)

def pdf_to_excel_3(csvname,xlsxname):
    merge_all_to_a_book(glob.glob("output/excel/*.csv"), "output.xlsx")
#    merge_all_to_a_book(output_excel_path+csvname, output_excel_path+xlsxname)


# pdf_to_excel_3(csvname,xlsxname)    


def pdf_to_excel_1(pdf_name,xlsxname):
    # Read PDF File
    df = tabula.read_pdf(input_agile_path+pdf_name, pages = 1)[0]
    # Convert into Excel File
    tabula.convert_into(input_agile_path+pdf_name, output_excel_path+csvname, output_format="csv", pages='all',stream=True)

pdf_to_excel_1(pdf_name,xlsxname)

def csv_to_excel_3(csvname,xlsxname):
    wb = Workbook()
    ws = wb.active
    with open("output/excel/agile_try.csv", 'r') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save('name.xlsx')

csv_to_excel_3(csvname,xlsxname)   