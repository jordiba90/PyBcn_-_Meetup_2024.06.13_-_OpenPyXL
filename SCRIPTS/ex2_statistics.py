import os; clear = lambda: os.system('cls'); clear();
import openpyxl as xl; #print("Python Library: ", xl.__name__, xl.__version__)
import pandas as pd; import scipy; import numpy as np;

################################################################################

path_in = r'C:\Users\jordi\Desktop\ARCHIVO\1 - DOCS\PyBcn - OpenPyXL\XLS\ex1_countries.xlsx'
path_out = r'C:\Users\jordi\Desktop\ARCHIVO\1 - DOCS\PyBcn - OpenPyXL\XLS\ex2_statistics.xlsx'

################################################################################

print("\n***************************************************************************")
print("DB: Northwind; TB: OrderID")
print("***************************************************************************\n")

df = pd.read_excel(path_in, sheet_name=None)

wb = xl.Workbook()
wb.remove(wb.active)

def add_statistics_sheet(wb, sheet_name, data, numeric_only=True):

    if numeric_only:
        data = data.select_dtypes(include=[np.number])

    statistics = {
        'mean': data.mean(), 'median': data.median(), 'std_dev': data.std(), 'variance': data.var(),
        'min': data.min(),'max': data.max(),
        '25%': data.quantile(0.25), '50%': data.quantile(0.5), '75%': data.quantile(0.75),
    }

    ws = wb.create_sheet(title=sheet_name)
    ws.append(['Variable', 'Mean', 'Median', 'Std Dev', 'Variance', 'Min', 'Max', '25%', '50%', '75%'])

    for column in data.columns:
        row = [column] + [statistics[stat][column] for stat in statistics]
        ws.append(row)

for sheet_name, data in df.items():
    add_statistics_sheet(wb, sheet_name + ' Stats', data)

wb.save(path_out)

print("\nOK 😊👍\n")

print("***************************************************************************\n")

################################################################################
