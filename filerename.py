# import os   
# flepath = "//srva090/lngcew$/General/06_M_PROJECT MANAGEMENT & CONTROL/70_PIM/90 Personal/Sangit/EFiles ISO"

# firstChar = 'L001-00000-MP-2343-'

# for pdffile in os.listdir(flepath):
#     x = pdffile[:-8]
#     print(x)
#     newName = firstChar + x + '-R0.pdf'
#     print(newName)
#     os.rename(pdffile,newName)
  
import os
import shutil
from glob import glob

filepath = r'C:\USERDATA\General Tasks\EFiles ISO'

for file in next(os.walk(filepath))[2]:
    f = glob(os.path.join(filepath,"*_0_00001.pdf"))[0]
    print(f)
    if f:
        os.rename(os.path.join(filepath, file), "L001-00000-MP-2343-"+file[:21]+"_R0.pdf")
        print(f)
