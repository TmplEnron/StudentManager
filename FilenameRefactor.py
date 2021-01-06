import os
import xlrd
import shutil
from xlutils.copy import copy
import tqdm

name_list_excel=r"F:\web前端技术\(2020-2021-1)-0809412108-web前端开发技术.xls"

homework_orgin_path=r"F:\web前端技术\实验报告"

homework_target_path=r"F:\web前端技术\实验报告1.0"



if not os.path.exists(homework_target_path):
    os.makedirs(homework_target_path)


homework_docs=os.listdir(homework_orgin_path)

read_book = xlrd.open_workbook(name_list_excel)

sheet = read_book.sheet_by_index(0)

write_book = copy(read_book)

unsubmit_students=[];

for rowNum in range(5,sheet.nrows):
    print("Total: %d, current: %d",sheet.nrows,rowNum)
    studentId = sheet.row(rowNum)[1].value
    studentName=sheet.row(rowNum)[2].value
    studentClass=sheet.row(rowNum)[4].value
    nstudentName=studentName
    if len(studentName)==2:
        nstudentName=studentName[0]+"   "+studentName[1];
    submit=False
    for homework_doc in homework_docs:
        if  studentName in homework_doc or studentId in homework_doc:
            p1=os.path.join(homework_orgin_path, homework_doc)
            new_file_name=studentId + '-' + nstudentName + "-" + studentClass +'班'+ os.path.splitext(homework_doc)[-1]
            p2=os.path.join(homework_target_path,new_file_name)
            shutil.copy(p1,p2)
            submit=True;
            break
    if not submit:
        unsubmit_students.append(nstudentName)


print(unsubmit_students)