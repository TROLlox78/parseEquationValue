from curses.ascii import isdigit
import decimal
from pickletools import read_stringnl_noescape_pair
from openpyxl import load_workbook
import pytexit
filename = "aa.xlsx"
workbook = load_workbook(filename=filename )
workbook_data = load_workbook(filename=filename, data_only=True )
sheet = workbook.active
datasheet = workbook_data.active


def find_cell(s:str):
    # finds the first cell in string; like A5 or BC52
    i=0
    out = ''
    digitphase = False
    first = True
    while i!=len(s):
        ch = s[i]
        if first:
            if not ch.isalpha() and not ch.isupper() :
                return ''
            else:
                first = False;
        if ch.isalpha() and ch.isupper() and not digitphase:
            out+=ch
        elif ch.isdigit() and len(out)>0:
            out+=ch
            digitphase = True
        elif digitphase and not ch.isdigit():
            return out
        else:
            return ''

        i+=1
    return out

from math import sqrt

def clean_formula(val):
    val= val.replace('^','**')
    val = val.replace("SQRT" , 'sqrt')
    val = val.replace("-0" , '-')
    return val

def extract_values(cell_value):
    # finds the equation and values in the cell and builds a string
    s = ''
    skip_chr = 0;
    if not cell_value:
        return "0"
    if isinstance(cell_value,(int,float)) :
        return str(cell_value)
    for i in range(1,len(cell_value)):
        if skip_chr>0:
            skip_chr-=1;
            continue;
        newcell = find_cell(cell_value[i:])
        if len(newcell) >0:
            values = clean_formula(extract_values(sheet[newcell].value))
            s+= str(eval(values))
            skip_chr = len(newcell)-1
        else:
            s+=cell_value[i]
    return clean_formula(s)

def shorten(s:str):
    i = 0
    for ch in s:
        i+=1
        if ch==".":
            i+=2
            return s[:i]
        

def float_killer(s, round_to):
    check_num = False;
    decimal_point = False;
    s = repr(s)

    if round_to==0:
        return s

    new_s = ''
    for ch in s:
        if ch.isdigit() and not check_num:
            check_num = True
            decimal_point = False
        if check_num:
            if ch.isdigit() and not decimal_point:
                new_s +=ch
            if ch == '.':
                count = 0
                decimal_point = True
                new_s+=ch
            if decimal_point and ch.isdigit():
                count+=1
                if count<round_to:
                    new_s +=ch
            if not ch.isdigit() and ch != '.':
                check_num = False;
                new_s+=ch
        else:
            new_s +=ch
    return new_s

def extract(cell):
    
    equation =extract_values(cell)
    latex = pytexit.py2tex(equation, print_latex=False, print_formula=False)
    return latex
        #print(latex + " = " + shorten(str(eval(equation))))





if __name__ == "__main__":
    cells = []
    search_col = 29+3
    for i in range(31,70):
        name = sheet.cell(row=i, column=search_col-1).value
        #print(sheet.cell(row=i, column=search_col).value[1:])
        latex = extract(sheet.cell(row=i, column=search_col).value)
        value = float_killer(datasheet.cell(row=i, column=search_col).value,0)
        print(f'{name} = {latex} = {value}')
