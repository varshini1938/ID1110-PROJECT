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
		# If threshold crossed, modify the message
		if l2 != "" and len(l3) != 0:

			# message for student
			msg1 = "You have lack of attendance in " + subject + " !!!"

			# message for staff
			msg2 = "The following students have lack of attendance in your subject : "+l2

			mailstu(l3, msg1) # mail to students
			staff_id = staff_mails[b-1] # pick respective staff's mail_id
			mailstaff(staff_id, msg2) # mail to staff
        

# for students
 def mailstu(li, msg):
	from_id = '132301012@smail.iitpkd.ac.in'
	pwd = 'Ts07fm4578!'
	s = smtplib.SMTP('smtp.gmail.com', 587, timeout=120)
	s.starttls()
	s.login(from_id, pwd)

	# for each student to warn send mail
	for i in range(0, len(li)):
		to_id = li[i]
		message = MIMEMultipart()
		message['Subject'] = 'Attendance report'
		message.attach(MIMEText(msg, 'plain'))
		content = message.as_string()
		s.sendmail(from_id, to_id, content)
	s.quit()
	print("mail sent to students")

# for staff
def mailstaff(mail_id, msg):
	from_id = '132301012@smail.iitpkd.ac.in'
	pwd = 'Ts07fm4578!'
	to_id = mail_id
	message = MIMEMultipart()
	message['Subject'] = 'Lack of attendance report'
	message.attach(MIMEText(msg, 'plain'))
	s = smtplib.SMTP('smtp.gmail.com', 587, timeout=120)
	s.starttls()
	s.login(from_id, pwd)
	content = message.as_string()
	s.sendmail(from_id, to_id, content)
	s.quit()
	print('Mail Sent to staff')

while resp == 1:
	print("1--->phy\n2--->math\n3--->mech")

	# enter the correspondingnumber
	y = int(input("enter subject :"))

	# no.of.absentees for that subject
	no_of_absentees = int(input('no.of.absentees :'))

	if(no_of_absentees > 1):
		x = list(map(int, (input('roll nos :').split(','))))
	else:
		x = [int(input('roll no :'))]

	# list to hold row of the student in Excel sheet
	row_num = []

	# list to hold total no.of leaves
	# taken by ith student
	no_of_days = []

	for student in x:

		for i in range(2, rows+2):

			if y==1:
				if sheet.cell(row=i, column=1).value == student:
					m = sheet.cell(row=i, column=3).value
					m = m+1
					sheet.cell(row=i, column=3).value = m
					savefile()
					no_of_days.append(m)
					row_num.append(i)

			elif y == 2:
				if sheet.cell(row=i, column=1).value == student:
					m = sheet.cell(row=i, column=4).value
					m = m+1
					sheet.cell(row=i, column=4).value = m
					no_of_days.append(m)
					row_num.append(i)

			elif y == 3:
				if sheet.cell(row=i, column=1).value == student:
					m = sheet.cell(row=i, column=5).value
					m = m+1
					sheet.cell(row=i, column=5).value = m
					row_num.append(i)
					no_of_days.append(m)

	check(no_of_days, row_num, y)
	resp = int(input('another subject ? 1---->yes 0--->no'))
	
		
