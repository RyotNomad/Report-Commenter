import os
import shutil
import sys
import errno
from docx import Document



def get_remark(mark):

    mark = int(mark)
    if(mark) > 80:
        return "Outstanding Achievement"
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
        return "Not Achieved"

backup_name = "Report_backups"
backup_name_dir = "/Report_backups"


try:
    os.mkdir(backup_name)
    print("Backup successfully made")
    sys.exit()

except OSError as e:
    if e.errno == errno.EEXIST:
        print('Backup already exists. Contact Farhaan')


files = [f for f in os.listdir('.') if os.path.isfile(f)]

#Create backups
for f in files:
    print(f)
    if f != "reportCommenter.py":
        shutil.copy(f,backup_name)

#Open files

for f in files:
    if f == "trt.docx":
        print("working...")
        doc = Document(f)
        for table in doc.tables:
            for i in range(1 ,len(table.rows)):
                mark = table.cell(i,2)

                comment = table.cell(i, 3)
                remark = get_remark(mark.text)
                table.cell(i,3).text = remark
            doc.save(f)
print("Saving...")




