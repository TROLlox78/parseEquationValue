from curses.ascii import isdigit
import decimal
from pickletools import read_stringnl_noescape_pair
from openpyxl import load_workbook
import pytexit
import re

# this import sys bad, hack for easy stdout redirect to file
import sys
sys.stdout.reconfigure(encoding='utf-8')
filename = "aue_lab7.xlsx"
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

# this is necessary
from math import sqrt

def clean_formula(val):
    # replaces excel to python lang
    val= val.replace('^','**')
    val = val.replace("SQRT" , 'sqrt')
    val = val.replace("-0" , '-')
    #val = val.replace("PI()" , '3.14')
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

        

def extract(cell):
    # creates latex 
    equation =extract_values(cell)
    latex = pytexit.py2tex(equation, print_latex=False, print_formula=False)
    return latex

def shorten_small_float(n : str, round_to):
    # run for float<1
    output = n[:2]
    count =0
    counting = False
    # cut out "0."" wtih [2:]
    for ch in n[2:]:

        output+=ch
        if int(ch)>0:
            # found nonzero digit
            counting=True
        if counting:
            count+=1
        if count == round_to:
            return output
    return output

def reduce_floats(s: str, round_to=2):
    find_floats = r"\d+[.]\d+"
    #teststr = repr("Tl1= = $$9.15\times{10}^{-12} \left(4600+292.7351842831629\times{10}^6\right)$$ = 0.0026785690261909405")

    def matchfloat(matchobj):
        round_to = 2
        match = float(matchobj.group(0))
        if match<1:
            # so that you don't turn 0.004 => 0.0
            return shorten_small_float(str(match), round_to)

        return str(round(match, round_to))
    
    value_short = re.sub(find_floats, matchfloat ,s)
    return value_short


if __name__ == "__main__":
    cells = []

    # set cell search scope
    start_cell = 22
    end_cell   = 40
    value_letter = 'u'


    search_col = int(ord(value_letter))-int(ord('a'))+1
    for i in range(start_cell,end_cell):
        cell_name = sheet.cell(row=i, column=search_col-1).value
        #print(sheet.cell(row=i, column=search_col).value[1:])
        latex = extract(sheet.cell(row=i, column=search_col).value)
        value = datasheet.cell(row=i, column=search_col).value

        output = reduce_floats(f'{cell_name} = {latex} = {value}',2)
        print(f'{cell_name} = {latex} = {value}')
        print(output)
