# -*- coding: utf-8 -*-

import os
import sys
import csv
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side

class MultipleCSVToExcel:
    DEST_ROW_INDEX = 2
    DEST_COLUMN_INDEX = 2
    TEMPLATE_PATH = os.path.dirname(os.path.abspath(__file__)) + "\\template"
    CSV_PATH = os.path.dirname(os.path.abspath(__file__)) + "\\csv"
    TEMPLATE_FILE = "template.xlsx"
    NEW_FILE = "new_file.xlsx"
    SEPARATE = "." 
    FONT = Font(name = "游ゴシック Medium", size = 11)
    SIDE = Side(border_style = "thin", color = "000000")
    BORDER = Border(top = SIDE, bottom = SIDE, left = SIDE, right = SIDE)

    def __paste_data(self, worksheet, data):
        for row_index, row_value in enumerate(data):
            for column_index in range(len(row_value)):
                cell = worksheet.cell(row = self.DEST_ROW_INDEX + row_index, \
                                        column = self.DEST_COLUMN_INDEX + column_index)
                cell.value = row_value[column_index]
                cell.font = self.FONT
                cell.border = self.BORDER


    def __open_csv_file(self, worksheet, csv_file):
        with open(self.CSV_PATH + "\\" + csv_file, "r") as f:
            csv_data = csv.reader(f, delimiter=",")
            self.__paste_data(worksheet, csv_data)


    def __copy_csv_to_excel(self, workbook, csv_files):
        worksheet = workbook.active
        for csv_file in csv_files:
            copy_worksheet = workbook.copy_worksheet(worksheet)
            copy_worksheet.title = csv_file.split(self.SEPARATE, 1)[0]
            self.__open_csv_file(copy_worksheet, csv_file)
        workbook.save(self.NEW_FILE)


    def __open_workbook(self):
        return load_workbook(self.TEMPLATE_PATH + "\\" + self.TEMPLATE_FILE)


    def __get_csv_files(self):
        csv_files = os.listdir(self.CSV_PATH)
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
        csv_files = self.__get_csv_files()
        wb = self.__open_workbook()
        self.__copy_csv_to_excel(wb, csv_files)


if __name__ == "__main__":
    MultipleCSVToExcel().main()