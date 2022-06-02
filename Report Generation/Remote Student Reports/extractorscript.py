import pandas as pd
import xlsxwriter

import smtplib, ssl

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import getpass

smtp_server = "smtp.gmail.com"

port = 587

sender_email = "noreplyPELLearning@gmail.com"

passcode = getpass.getpass("Enter the autoemailing passcode ")


context = ssl.create_default_context()

server = None

try:
	server = smtplib.SMTP(smtp_server, port)
	server.starttls(context = context)
	server.login(sender_email, passcode)

except Exception as e:
	print(e)



month = 1

current_students = 0


line_chart = [[], [], [], [], []]


while (True):
	new_students = 0
	leaving_students = 0
	skipping_students = 0
	try:
		data = pd.read_excel("data.xlsx", index_col="No.", sheet_name = f"Month{month}")

	except ValueError:
		break

	for ind in range(1,len(data)+1):
		if data.loc[ind, "Code (C, M, N, A, R)"] == 'N':
			new_students+=1
		elif data.loc[ind, "Code (C, M, N, A, R)"] == 'A':
			skipping_students+=1


	leaving_students = current_students - (len(data) - new_students)
	

	
	current_students = len(data) - skipping_students

	if (month == 1):
		current_students = len(data)


	else:


		line_chart[0].append(month)
		line_chart[1].append(current_students)
		line_chart[2].append(new_students)
		line_chart[3].append(leaving_students)
		line_chart[4].append(skipping_students)


	current_students = len(data)

	

	month += 1


month -= 1



# This Month's Charts


workbook = xlsxwriter.Workbook(f'Report{month}.xlsx')
worksheet = workbook.add_worksheet(f"Month{month}")

#Line Chart

linechart = workbook.add_chart({'type': 'line'})

worksheet.write_row('E1', ['Month', 'Current Students', 'New Students', 'Students Leaving', 'Absent Students'])
worksheet.write_column('E2', line_chart[0])
worksheet.write_column('F2', line_chart[1])
worksheet.write_column('G2', line_chart[2])
worksheet.write_column('H2', line_chart[3])
worksheet.write_column('I2', line_chart[4])

for column in ['F', 'G', 'H', 'I']:

	linechart.add_series({
		'name' : f'=Month{month}!${column}$1',
		'categories' : f'=Month{month}!$E$2:$E${month}',
		'values': f'=Month{month}!${column}$2:${column}${month}'
		
		})


linechart.set_title({'name' : 'Program Popularity'})
linechart.set_x_axis({'name' : 'Month'})
linechart.set_y_axis({'name' : '# of Students'})

# Bar Chart


data = pd.read_excel("data.xlsx", index_col="No.", sheet_name = f"Month{month}")


chart = workbook.add_chart({'type' : 'column'})

englishByGrade = [0 for i in range(1,14)];
englishByGrade[0] = 'Reports'
mathByGrade = [0 for i in range(1,14)];
mathByGrade[0] = 'Reports'


kinm = 0
kine = 0

for ind in range(1,len(data)+1):
	if (data.loc[ind, 'School Grade'] == 'K'):
		kinm += 1 if data.loc[ind, 'Subject (M/E)'] == 'M' else 0
		kine += 1 if data.loc[ind, 'Subject (M/E)'] == 'E' else 0
	elif data.loc[ind, 'Subject (M/E)'] == 'M':
		englishByGrade[data.loc[ind, 'School Grade'] + 2] += 1
	elif data.loc[ind, 'Subject (M/E)'] == 'E':
		mathByGrade[data.loc[ind, 'School Grade'] + 2] += 1

	else:
		print("Invalid Grade Input in row " + ind)




worksheet.write('A1', 'Grade Level')

worksheet.write('B1', 'English')
worksheet.write('C1', 'Math')

worksheet.write_row('A2', ['K', kine, kinm])

for i in range(3, 15):
	worksheet.write('B' + str(i), englishByGrade[i-2])
	worksheet.write('A' + str(i),i-2)
	worksheet.write('C' + str(i), mathByGrade[i-2])


chart.add_series({
	'categories': [f'Month{month}', 1, 0, 13, 0],
	'values': [f'Month{month}', 1, 1, 13, 1],
	'fill': {'color': 'red'},
	'name' : f'=Month{month}!$B$1'
	})

chart.add_series({
	'categories': [f'Month{month}', 1, 0, 13, 0],
	'values': [f'Month{month}', 1, 2, 13, 2],
	'fill': {'color': 'blue'},
	'name' : f'=Month{month}!$C$1'
	})
# chart.add_series({'values': mathByGrade})

chart.set_x_axis({
	'name' : f'=Month{month}!$A$1',
	'name_font': {'size':14, 'bold':True},
	'num_font' : {'italic':True}

	})



chart.set_y_axis({
	'name' : 'Number of Students',
	'name_font': {'size':14, 'bold':True},
	'num_font' : {'italic':True}
	})


chart.set_title({'name' : 'Enrollment by Grade Level'})

worksheet.insert_chart("K1", chart)
worksheet.insert_chart("K17", linechart)
workbook.close()


# Send Emails



message = MIMEMultipart("alternative")

message["Subject"] = "Thanks from PEL Learning"
message["From"] = sender_email
message["To"] = data.loc[ind, 'Email']
if data.loc[ind, 'Code (C, M, N, A, R)'] == 'R':
	msg = "Thanks for continuing " + data.loc[ind, "First Name"] + ' ' + data.loc[ind, "Last Name"] + " with PEL learning"
elif data.loc[ind, 'Code (C, M, N, A, R)'] == 'A':
	msg = "Could you please tell us why " + data.loc[ind, "First Name"] + ' ' + data.loc[ind, "Last Name"] + " is not continuing with PEL learning? It will help us improve our program!"
elif data.loc[ind, 'Code (C, M, N, A, R)'] == 'N':
	msg = "Thanks for enrolling " + data.loc[ind, "First Name"] + ' ' + data.loc[ind, "Last Name"] + " in PEL learning's program!"





msg = MIMEText(msg, 'plain')

message.attach(msg)

server.sendmail(sender_email, data.loc[ind, 'Email'], message.as_string())

server.quit()
