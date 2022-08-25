import re
import pandas as pd


def get_input():
    loc_write = ""
    loc_read = ""

    readfile1 = False
    readfile2 = False

    while not readfile1:
        loc_write = input("\033[94mPlease submit UTL file:\033[0m ").replace("\\", "").replace("\'", "").rstrip()
        try:
            pd.read_excel(loc_write)
            readfile1 = True
        except:
            print("\033[1;91mCan't read file, please insert correct file\033[0m")

    while not readfile2:
        loc_read = input("\033[35mPlease submit Korea Comparison:\033[0m ").replace("\\", "").replace("\'", "").rstrip()
        try:
            pd.read_excel(loc_read)
            readfile2 = True
        except:
            print("\033[1;91mCan't read file, please insert correct file\033[0m")

    return loc_write, loc_read


def get_maximum_rows(*, sheet_object):
    rows = 0
    for max_row, row in enumerate(sheet_object, 1):
        if not all(col.value is None for col in row):
            rows += 1
    return rows


def get_start_rows(sheet_object):
    col_num = find_match_key(sheet_object[1], '^inv')
    for key, row in enumerate(sheet_object, -1):
        if row[col_num].value is None:
            return key

    return 0


def find_match_val(df, reg):
    for col in df.columns:
        if re.search(reg, col, flags=re.IGNORECASE):
            return col


def find_match_key(sheet_val, reg):
    for k, v in enumerate(sheet_val):
        if re.search(reg, v.value, flags=re.IGNORECASE):
            return k


def rename_column(df, reg_list, name_list):
    for i in range(len(name_list)):
        df = df.rename(columns={find_match_val(df, reg_list[i]): name_list[i]})
    return df
