from openpyxl import load_workbook
import pytexit
workbook = load_workbook(filename="aa.xlsx" )
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
            continue;
        newcell = find_cell(cell_value[i:])
        if len(newcell) >0:
            s+= extract_values(newcell)
            skip_chr = len(newcell)-1
        else:
            s+=cell_value[i]
    return s

def shorten(s:str):
    i = 0
    for ch in s:
        i+=1
        if ch==".":
            i+=2
            return s[:i]
        

def extract(cells):
    for cell in cells:
        equation =extract_values(cell)
        print(cell)
        latex = pytexit.py2tex(equation, print_latex=False)
        #print(latex + " = " + shorten(str(eval(equation))))



cells = []

for i in range(31,54):
    if i==35 or i==36 or i ==47:
        continue
    cells.append(str('S'+str(i)))

if __name__ == "__main__":
    extract(cells)