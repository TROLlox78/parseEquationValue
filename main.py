from openpyxl import load_workbook
import sympy
import pytexit
workbook = load_workbook(filename="C:\\Users\\rectangle_man\\dev\\python\\dupa.xlsx" )
sheet = workbook.active


def find_cell(s:str):
    i=0
    while i!=len(s):
        ch = s[i]
        if ch.isalpha() and ch.isupper():
            print(ch)
        i+=1
    return True
find_cell("AstNerAL")

def extract_values(cell):
    s = ''
    skip_nxt_chr = False;

    cell_value = sheet[cell].value
    
    if isinstance(cell_value,(int,float)) :
        return str(cell_value)
    for i in range(1,len(cell_value)-1):
        if skip_nxt_chr:
            skip_nxt_chr =False
            continue;
        if cell_value[i].isalpha() and cell_value[i+1].isdigit():
            newcell = cell_value[i]+cell_value[i+1]
            s+= extract_values(newcell)
            skip_nxt_chr = True
        else:
            s+=cell_value[i]
    return s

def extract():
    x =extract_values("A1")
    print(x)
    y = sympy.sympify(x)
    pytexit.py2tex(x)
    print(sympy.latex(y))
