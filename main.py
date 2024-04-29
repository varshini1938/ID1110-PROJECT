#staff mail ids
staff_mails=['srivarshini4578@gmail.com','132301012@smail.iitpkd.ac.in','132301012@smail.iitpkd.ac.in']

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

























