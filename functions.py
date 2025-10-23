import pandas as pd
import openpyxl
from openpyxl.styles import  Font, Border, Side, Alignment, numbers 
from openpyxl import load_workbook,Workbook

def write_summary_formula(ws, min_row, min_col, max_row, max_col,sum_startrow,sum_endrow):
# function to dynamically write sum formula
# min_row: row where sum should be placed
# min_col: column where sum should be placed
# sum_startrow: row from where sum calculation should be starting
# sum_endrow: row from where sum calculation should be ending
    for col in ws.iter_cols(min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col): # range where to start and end loop
        for cell in col:
            cell_sum_start = cell.column_letter + str(sum_startrow)
            cell_sum_end = cell.column_letter + str(sum_endrow)
            cell.value = f'=SUM({cell_sum_start}:{cell_sum_end})'
            cell.number_format = '_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)' # Acconting stype without $. Source: https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/numbers.html
            cell.font = Font(bold=True)
            #cell.border = double_border


# create double border for summary row
def apply_double_border_range(ws, start_row, start_column, end_row, end_column, double_border):
    # Iterate over the range of cells and apply double border
    for row in range(start_row, end_row + 1):
        for column in range(start_column, end_column + 1):
            ws.cell(row=row, column=column).border = double_border

# top and double buttom border style
double_border_format = Border(top=Side(style='thin'), bottom=Side(style='double'))


# adjust column widts. Pass multiple columns Ex: ['A', 'B', 'C']
def adjust_columns_width(ws, columns_to_adjust, width):
    for col in columns_to_adjust:
        ws.column_dimensions[col].width = width


def header_formatting(ws, df):
    # write Carrier name 
    carrier = ws.cell(row=1, column=1)
    carrier.value = 'AXIS Insurance Company'
    carrier.font = Font(name='Arial', bold=True, size = "10")

    # write office name
    office = ws.cell(row=2, column=1)
    office.value = 'DUAL North America'
    office.font = Font(name='Arial',bold=True, size = "10")

    # write program name
    program = ws.cell(row=3, column=1)
    program.value = 'Surety Bond Program'
    program.font = Font(name='Arial',bold=True, size = "10")

    # write prod report
    prodreport = ws.cell(row=4, column=1)
    prodreport.value = 'Production Report'
    prodreport.font = Font(name='Arial', bold=True, size = "10")

    # grab top 1 DateFrom and DateTo 
    dateFrom = df['DateFrom'].iloc[0]
    dateTo = df['DateTo'].iloc[0]
    # Format dates in mm/dd/yyyy format
    dateFrom_str = dateFrom.strftime('%m/%d/%Y')
    dateTo_str = dateTo.strftime('%m/%d/%Y')
    dateFromTo = f"For the period {dateFrom_str} to {dateTo_str}"
    ws['A5'].value = dateFromTo
    ws['A5'].font = Font(bold=True)


    # since iterating through each row of df individually - its ok to grab top 1 "PaymentDue"
    paymentDue = df['PaymentDue'].iloc[0]
    paymentTermsDue_str = paymentDue.strftime('%m/%d/%Y')
    paymentTermsDue = f"Payment Terms Due: {paymentTermsDue_str}"
    ws['A6'].value = paymentTermsDue
    ws['A6'].font = Font(bold=True)



def replace_zeros_in_columns(ws, columns_to_replace):
    """
    Replace 0 with '-' in specified columns of a given worksheet.

    Parameters:
    ws (openpyxl.worksheet.worksheet.Worksheet): The worksheet object to modify.
    columns_to_replace (list): List of column letters to replace 0 with '-'.
    """
    # Iterate through the rows in the sheet
    for row in ws.iter_rows():
        for col_letter in columns_to_replace:
            cell = row[openpyxl.utils.cell.column_index_from_string(col_letter) - 1]
            if cell.value == 0:
                cell.value = '-'
                cell.alignment = Alignment(horizontal='right')


