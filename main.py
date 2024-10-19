from openpyxl import load_workbook
import sympy
import pytexit
workbook = load_workbook(filename="dupa.xlsx" )
#workbook = load_workbook(filename="C:\\Users\\rectangle_man\\dev\\python\\parseEquationValue\\dupa.xlsx" )
sheet = workbook.active


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

        i+=1
    return out


def extract_values(cell):
    # finds the equation and values in the cell and builds a string
    s = ''
    skip_chr = 0;

    cell_value = sheet[cell].value
    
    if isinstance(cell_value,(int,float)) :
        return str(cell_value)
    for i in range(1,len(cell_value)):
        if skip_chr>0:
            skip_chr-=1;
            #print("skip:"+cell_value[i])
            continue;
        newcell = find_cell(cell_value[i:])
        if len(newcell) >0:
            print(newcell)
            s+= extract_values(newcell)
            skip_chr = len(newcell)-1
        else:
            s+=cell_value[i]
            print(cell_value[i])
    return s

def extract():
    x =extract_values("A2")
    print(x)
    #y = sympy.sympify(x)
    #pytexit.py2tex(x)
    #print(sympy.latex(y))
extract()