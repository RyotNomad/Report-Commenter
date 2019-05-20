import os
import shutil
import sys
from docx import Document



def get_remark(mark):
    mark = int(mark)
    if(mark) > 70:
        return "Meritorious Achievement"
    if(mark>60):
        return "Substantial Achievement"
    if(mark>50):
        return "Adequate Achievement"
    if(mark>40):
        return "Moderate Achievement"
    if(mark>30):
        return "Elementary Achievement"
    else:
        return "Hella dom or hella smart"

backup_name = "Report_backups"
backup_name_dir = "/Report_backups"


try:
    os.mkdir(backup_name)
    print("Backup successfully made")

except FileExistsError:
    print("Error it seems like the backup has already been made, please contact Farhaan")
    #sys.exit()

files = [f for f in os.listdir('.') if os.path.isfile(f)]

#Create backups
for f in files:
    print(f)
    if f != "reportCommenter.py":
        shutil.copy(f,backup_name)

#Open files

for f in files[1:len(files)-2]:
    doc = Document(f)
    for table in doc.tables:
        for i in range(0,4):
            mark = table.cell(i, 0)
            comment = table.cell(i, 1)
            remark = get_remark(mark.text)
            table.cell(i,1).text = remark
        doc.save(f)



