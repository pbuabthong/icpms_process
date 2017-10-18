import openpyxl as opx
import numpy as np
from scipy import interpolate

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def findFirstCell(sheet, search_str):
    for row in sheet.iter_rows():
        for cell in row:
            if search_str in str(cell.value):
                # print("Final: found at", opx.utils.cell.coordinate_to_tuple(cell.coordinate))
                return opx.utils.cell.coordinate_to_tuple(cell.coordinate)

raw_filename = 'raw.xlsx'
std_filename = 'std.xlsx'
alz_filename = 'alz.xlsx'
el = 'Te'

#wb = opx.load_workbook('raw.xlsx')
print("Loading standard file:", std_filename)
std = opx.load_workbook(std_filename, data_only = True)
std_sheet = std.get_sheet_by_name(std.get_sheet_names()[0])
keyword_coord = findFirstCell(std_sheet,'Final:')
element_list_row = keyword_coord[0]-1

std_list = []
for element_col in std_sheet.iter_cols(min_row = element_list_row, min_col = (keyword_coord[1]+1)):
    if el in element_col[0].value:
        for cell in element_col[2:]:
            std_list.append(cell.value)
std_list = std_list[::-1]
# print(std_list)

raw = opx.load_workbook(raw_filename, data_only = True)
raw_sheet = raw.get_sheet_by_name(raw.get_sheet_names()[0])
# raw_keyword_coord = findFirstCell(raw_sheet, 'Te')
# print(raw_keyword_coord)

raw_list = []
for raw_element_col in raw_sheet.iter_cols(max_row = 13):
    if el in str(raw_element_col[0].value):
        for cell in raw_element_col[3:]:
            raw_list.append(cell.value)

# print(raw_list)

# f1d = interpolate.interp1d(raw_list, std_list)
f1d = interpolate.InterpolatedUnivariateSpline(raw_list, std_list, k=1)

alz = opx.load_workbook(alz_filename, data_only = True)
alz_sheet = alz.get_sheet_by_name(alz.get_sheet_names()[0])

samplename_column = findFirstCell(alz_sheet, 'Sample Name')[0]

name_list = []
# rsd = False
for raw_element_col in raw_sheet.iter_cols():
    if 'Sample Name' in str(raw_element_col[1].value):
        for cell in raw_element_col[2:]:
            name_list.append(cell.value)

    # if (el in str(raw_element_col[0].value)) or (rsd == True):
    # TODO: not sure how rsd values are calculated
    if (el in str(raw_element_col[0].value)):
        i = 0
        # status = 'RSD' if rsd else 'CPS'
        status = 'CPS'
        for cell in raw_element_col[2:]:
            if is_number(cell.value):
                try:
                    current_coord = cell.coordinate
                    value_new = float(f1d(cell.value))
                    value_old = alz_sheet[current_coord].value
                    # print(value_new, ":", cell.value)
                    # print(type(value_new), ":", type(cell.value))
                    print("Converting", name_list[i], status, ": (", cell.value, " )", value_old, "-->", value_new)
                    alz_sheet[current_coord].value = value_new
                except:
                    print(name_list[i], ": cannot convert -- value: ", cell.value)
            else:
                print(name_list[i], ": not float --  value: ", cell.value)
            i += 1
        # rsd = not rsd

alz.save('alz2.xlsx')
