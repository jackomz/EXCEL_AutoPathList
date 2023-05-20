# -*- encoding: iso-8859-15 -*-

from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment  
from pathlib import Path
import tkinter
import tkinter.filedialog
import os


KiB = 1024
MiB = KiB * KiB
GiB = MiB * KiB

root = tkinter.Tk()
root.withdraw()         #use to hide tkinter window
tempdir = tkinter.filedialog.askdirectory(parent=root, initialdir=os.getcwd(), title='Please select a directory')
root_directory = Path(tempdir)

Dateiname = "Liste.xlsx"
wb = Workbook()
ws = wb.active

ws['A1'] = "Path"
ws['B1'] = "Name"
ws['C1'] = "Size in GiB"
ws['D1'] = "Size in GB"


ListPath = tempdir + "/" + Dateiname

if len(tempdir) > 0:
    print ("You chose %s" % tempdir)

i = 2
for subFolder in os.listdir(tempdir):
    t = subFolder
    subFolder = tempdir+ "/" + subFolder
    temp = sum(f.stat().st_size for f in Path(subFolder).glob('**/*') if f.is_file())    
    ws['A'+str(i)] = subFolder
    ws['B'+str(i)] = str(t)
    ws['C'+str(i)] = round(temp/GiB,3)
    ws['D'+str(i)] = round(temp/10**9,3)
    i += 1


SumCell = "A"+str(i+1)+":B"+str(i+1)
SumCellGiB = "C"+str(i+1)
SumCellGB = "D"+str(i+1)

ws['A'+str(i+1)] = "Sum: "
ws[SumCellGiB] = "=Sum(C2:C"+str(i-1)+")"
ws[SumCellGB] = "=Sum(D2:D"+str(i-1)+")"
ws.merge_cells(SumCell)
    
wb.save(ListPath)
StartEXCELFile = "start EXCEL.EXE " + ListPath
os.system(StartEXCELFile)
print("List got generated! Location: \n\t" + ListPath)




