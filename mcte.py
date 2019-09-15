import os
import sys
import csv
from openpyxl import load_workbook

class MultipleCSVToExcel:

    def __init__(self):
        self.DEST_ROW_INDEX = 2
        self.DEST_COLUMN_INDEX = 2
        self.template_path = os.path.dirname(os.path.abspath(__file__)) + "\\template"
        self.csv_path = os.path.dirname(os.path.abspath(__file__)) + "\\csv"
        self.template_file = "template.xlsx"
        self.new_file = "new_file.xlsx"
        self.separate = "."     

    def copy_worksheet(self, workbook, csv_files):
        worksheet = workbook.active
        copy_worksheet = workbook.copy_worksheet(worksheet)
        copy_worksheet.title = ""


    def paste_data(self, worksheet, data):
        for row_index, row_value in enumerate(data):
            for column_index in range(len(row_value)):
                worksheet.cell(row = self.DEST_ROW_INDEX + row_index, \
                            column = self.DEST_COLUMN_INDEX + column_index, \
                            value = row_value[column_index])
                
    
    def open_csv_file(self, worksheet, csv_file):
        with open(self.csv_path + "\\" + csv_file, "r") as f:
            csv_data = csv.reader(f, delimiter=",")
            self.paste_data(worksheet, csv_data)


    def copy_csv_to_excel(self, workbook, csv_files):
        worksheet = workbook.active
        for csv_file in csv_files:
            copy_worksheet = workbook.copy_worksheet(worksheet)
            copy_worksheet.title = csv_file.split(self.separate, 1)[0]
            self.open_csv_file(copy_worksheet, csv_file)
        
        workbook.save(self.new_file)
            

    def open_workbook(self):
        return load_workbook(self.template_path + "\\" + self.template_file)


    def get_csv_files(self):
        csv_files = os.listdir(self.csv_path)
        if not csv_files:
            print("ERROR: Nothing csv files in csv direcoty. please store csv file.")
            sys.exit()

        return csv_files         
        

    def parser(self):
        pass
        ## 以下を参考に作る。sys.argvを使う
        ## https://qiita.com/petitviolet/items/aad73a24f41315f78ee4
        ## -r, --row integer
        ## -c, --column integer
        ## -d, --delimiter string

    def main(self):
        csv_files = self.get_csv_files()
        wb = self.open_workbook()
        self.copy_csv_to_excel(wb, csv_files)
        #copy_ws = self.copy_worksheet(wb, csv_files)


if __name__ == "__main__":
    mcte = MultipleCSVToExcel()
    mcte.main()