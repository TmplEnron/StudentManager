import os
import xlrd

group_excel=r"G:\大数据基础大作业\大作业\大数据开发基础大作业分组.xlsx"
experimental_excel=r"G:\大数据基础大作业\2020-2021-1-大数据开发基础名单.xls"

g_read_book = xlrd.open_workbook(group_excel)

g_sheet = g_read_book.sheet_by_index(0)

e_read_book = xlrd.open_workbook(experimental_excel)

e_sheet = e_read_book.sheet_by_index(0)

experimental_grades={}

group_grades={}
group_grade_details={}

student_to_group={}


for rowNum in range(5,e_sheet.nrows):
    student_name= e_sheet.row(rowNum)[2].value
    grade = e_sheet.row(rowNum)[14].value
    experimental_grades[student_name.strip()]=grade


for rowNum in range(1,g_sheet.nrows):
    print("Total: %d, current: %d",g_sheet.nrows,rowNum)
    gId = g_sheet.row(rowNum)[0].value
    if gId is None :
        break
    col=1
    grade=0
    count=0
    while True:
        if col>=g_sheet.ncols:
            break
        student_name = g_sheet.row(rowNum)[col].value
        if student_name is None or student_name.strip()=="":
            break
        student_name=student_name.strip()
        if student_name in experimental_grades.keys():
            grade=grade+experimental_grades[student_name]
            count=count+1
            student_to_group[student_name]=gId
        else:
            print(student_name)
        col = col + 1
    group_grades[gId]=round(grade/count,0)
    group_grade_details[gId] = [grade, count]
print(group_grade_details)
for key in group_grades.keys():
    print(group_grades[key])

for rowNum in range(5,e_sheet.nrows):
    student_name= e_sheet.row(rowNum)[2].value
    gId=student_to_group[student_name.strip()]
    print(student_name.strip(),'\t',gId,'\t',group_grades[gId])