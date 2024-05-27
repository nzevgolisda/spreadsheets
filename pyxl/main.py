
from openpyxl import Workbook
from openpyxl.styles import Alignment

def fillRow(i, student_data):
    header_data = ['Names', 'AM', 'Mail', 'Tests', 'Exams', 'Final']
    letters = list('ABCDEFGH')
    for j in range(len(header_data)):
        col = letters[j+1] + str(1)
        sheet[col] = header_data[j]
    sheet['A1'] = 'PK' #Primary key column
    for k in range(len(student_data[0])):
        col = letters[0]+str(k+2)
        sheet[col] = k+1
    letters = list('ABCDEFG')
    for j in range(1, len(letters)):
        col = letters[j] + str(i + 2)
        if j == len(letters) - 1:
            sheet[col] = student_data[3][i] * 0.4 + student_data[4][i] * 0.6
        else:
            sheet[col] = student_data[j-1][i]
Names = [
    'Nikos',
    'Tolis',
    'Adreas',
    'Maria'
]
AM = [4586, 4223, 5184, 7935]
Mail = [
    'nikos@gmail.com',
    'tolis@hotmail.com',
    'adreas@gmail.com',
    'maria@gmail.com'
]
Grades = [1, 8, 5, 2]
Final = [6, 7, 5, 7]
student_data = [Names, AM, Mail, Grades, Final]

workbook = Workbook()
sheet = workbook.active
for i in range(len(student_data[0])):
    fillRow(i, student_data)
header_data = ['Names', 'AM', 'Mail', 'Tests', 'Exams', 'Final']
rows = range(1, len(student_data[0])+2)
columns = range(1, len(header_data)+2)
for row in rows:
    for col in columns:
        Cell = sheet.cell(row, col)
        Cell.alignment = Alignment(horizontal = 'center', vertical = 'center')
widths = [6, 10, 5, 30, 6, 7, 5]
for col_letter, width in zip("ABCDEFG", widths):
    sheet.column_dimensions[col_letter].width = width
workbook.save(filename="./data.xlsx")
