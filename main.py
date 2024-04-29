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

# variable for looping for input
resp = 1

#updating the attendance to the excel sheet
while resp == 1:
    #subject numbers
    print("1--->phy\n2--->math\n3--->mech\n4--->python")

    # enter the corresponding number
    t = int(input("enter subject :"))

    # taking input of the number of people absent for that course
    no_of_absentees = int(input('no.of.absentees :'))

    #taking the roll no.'s of the students on leave 
    if no_of_absentees > 1:
        x = list(map(int, (input('roll nos :').split(','))))
    else:
        x = [int(input('roll no :'))]

    # list to hold row of the student in the excel sheet
    row_num = []

    # list to hold the total number of leaves taken by a particular student
    leaves = []

    #updating the excel sheet
    for n in x:
        #students
        for i in range(2, rows+2):
            if t == 1:
                if sheet.cell(row=i, column=1).value == n:
                    #updating the number of leaves
                    s = sheet.cell(row=i, column=3).value
                    s = s + 1
                    sheet.cell(row=i, column=3).value = s
                    #saving the data
                    savefile()
                    leaves.append(s)
                    row_num.append(i)
            elif t == 2:
                if sheet.cell(row=i, column=1).value == n:
                    s = sheet.cell(row=i, column=4).value
                    s = s + 1
                    sheet.cell(row=i, column=4).value = s
                    savefile()
                    leaves.append(s)
                    row_num.append(i)
            elif t == 3:
                if sheet.cell(row=i, column=1).value == n:
                    s = sheet.cell(row=i, column=5).value
                    s = s + 1
                    sheet.cell(row=i, column=5).value = s
                    savefile()
                    leaves.append(s)
                    row_num.append(i)
            elif t == 4:
                if sheet.cell(row=i, column=1).value == n:
                    s = sheet.cell(row=i, column=6).value
                    s = s + 1
                    sheet.cell(row=i, column=6).value = s
                    savefile()
                    leaves.append(s)
                    row_num.append(i)

    check(leaves, row_num, t)
    #taking the input if the user wants to check for another subject
    resp = int(input('another subject ? 1---->yes 0--->no'))

