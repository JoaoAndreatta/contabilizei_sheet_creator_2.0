import tkinter
import customtkinter
import openpyxl as xl

class Emprestimo():

    def test_function(self):
        self.label_test = customtkinter.CTkLabel(self.main_frame, text='Teste')
        self.label_test.grid(row=0,column=0)

    def create_loan_sheet(self):
        # creating workbook and worksheet
        wb = xl.Workbook()
        ws = wb.active
        ws.title = 'Depreciação'

        # definig columns names
        ws['A1'] = 'DATA'
        ws['B1'] = 'DESCRIÇÃO'
        ws['C1'] = 'VALOR'
        ws['D1'] = 'CONTA DÉBITO'
        ws['E1'] = 'CONTA CRÉDITO'

        # defining columns widths
        ws.column_dimensions['A'].width = 3
        ws.column_dimensions['B'].width = 36
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 15