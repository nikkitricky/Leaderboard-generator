import xlrd
import xlsxwriter

from xlrd import open_workbook

#This is the response excel sheet of google forms.. give correct path
submits = open_workbook('C:\\Users\\pc\\Desktop\CSI\\submissions.xlsx')

#This is the key sheet
answers = open_workbook('C:\\Users\\pc\\Desktop\CSI\\answers.xlsx')

submit_sheet = submits.sheet_by_index(0)
answers_sheet = answers.sheet_by_index(0)

scores = [] #scores of students
stud_ans = [] #answers of students

solutions = []

roll = []
for row in range(answers_sheet.nrows):
    solutions.append(str(answers_sheet.cell(row,0).value))


for row in range(1,submit_sheet.nrows):
    roll.append(str(submit_sheet.cell(row,1).value))


for row in range(1,submit_sheet.nrows):
    stud_ans = []
    for col in range(2,submit_sheet.ncols):
        if(submit_sheet.cell(row,col).value == xlrd.empty_cell.value):
            stud_ans.append('nikhith')
        else:
            stud_ans.append(str(submit_sheet.cell(row,col).value))

    #compare the answers
    marks = 0
    for i in range(0,38):
        if(solutions[i] == stud_ans[i]):
            marks += 1
    scores.append(marks)

workbook = xlsxwriter.Workbook('results.xlsx')
output=workbook.add_worksheet();
output.set_column('A:A',13)

for i in range(0,submit_sheet.nrows-1):
    output.write(i,0,roll[i])
    output.write(i,1,scores[i])

workbook.close()


