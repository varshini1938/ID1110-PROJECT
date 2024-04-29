import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# loading the excel sheet
book=openpyxl.load_workbook('C:/Users/smtca/Pictures/Sri/ACADEMICS/Python/attendance.xlsx')

# Choose the sheet
sheet=book['Sheet1']

# counting number of rows / students
rows=sheet.max_row-1

# variable for looping for input
resp=1

# counting number of columns / subjects
sub=sheet.max_column-2

# list of students to remind
l1=[]

# to concatenate list of roll numbers with
# lack of attendance
l2=""

# list of roll numbers with lack of attendance
l3=[]
def check(no_of_days, row_num, b):

	# to use the globally declared lists and strings
	global staff_mails
	global l2
	global l3

	for student in range(0, len(row_num)):
		# if total no.of.leaves equals threshold
		if no_of_days[student] == 2:
			if b==1:

				# mail_id appending
				l1.append(sheet.cell(row=row_num[student], column=2).value)
				mailstu(l1, message1) # sending mail
			
			elif b==2:
				l1.append(sheet.cell(row=row_num[student], column=2).value)
				mailstu(l1, message2)
			
			else:
				l1.append(sheet.cell(row=row_num[student], column=2).value)
				mailstu(l1, message3)

		# if total.no.of.leaves > threshold
		elif no_of_days[student] > 2:
			if b==1:

				# adding roll no
				l2=l2+str(sheet.cell(row=row_num[student], column=1).value)

				# student mail_id appending
				l3.append(sheet.cell(row=row_num[student], column=2).value)
				subject = "Physics"# subject based on the code number

			elif b==2:
				l2=l2+str(sheet.cell(row=row_num[student], column=1).value)
				l3.append(sheet.cell(row=row_num[student], column=2).value)
				subject = "Math"

			else:
				l2=l2+str(sheet.cell(row=row_num[student], column=1).value)
				l3.append(sheet.cell(row=row_num[student], column=2).value)
				subject = "Mechanics"
        

		
