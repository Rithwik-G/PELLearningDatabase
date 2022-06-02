
import xlsxwriter
import numpy as np 
from sklearn.linear_model import LinearRegression
import mysql.connector
import pandas as pd

db = mysql.connector.connect()

db = mysql.connector.connect(
host = "192.168.86.139",
user = "remoteuser",
passwd = "12345",
database = "StudentData"
)

mycursor = db.cursor()

month = 0

with pd.ExcelWriter('StudentData.xlsx') as writer:
	while True:
		month += 1
		sheet = [['No.',	'Subject (M/E)',	'First Name',	'Tel.',	'Email',	'DOE (Date of Enrollment MM/DD/YY)',	'PEL Wks. Level',	'PEL Wks. No.',	'Notes',	'Code (C, M, N, A, R)',	'Date Quit']]
		try:
			mycursor.execute(f"SELECT * FROM StudentDataMonth{month}")

			for value in mycursor:
				sheet.append(value)

			this_sheet = pd.DataFrame(sheet, columns = ['No.',	'Subject (M/E)',	'First Name',	'Tel.',	'Email',	'DOE (Date of Enrollment MM/DD/YY)',	'PEL Wks. Level',	'PEL Wks. No.',	'Notes',	'Code (C, M, N, A, R)',	'Date Quit'])
			this_sheet.to_excel(writer, sheet_name = f'Month{month}', header= False, index = False)





		except:
			print("Found data until month " + str(month - 1))
			break



for x in mycursor:
	print(x)

# ----------------

month = 1

student = input("Which student's report should be generated? ")
subject = input("For which subject should the report be generated (type M or E)? ").upper()

linechart = [[], []]

started = False

while True:
	try:
		data = pd.read_excel("StudentData.xlsx", index_col="No.", sheet_name = f"Month{month}")
	
	except ValueError:
		break

	for ind in range(1, len(data) + 1):
		if data.loc[ind, 'First Name'] == student and data.loc[ind, 'Subject (M/E)'] == subject:
			
			linechart[1].append(int(data.loc[ind, 'PEL Wks. No.'][:-1]))
			started = True

	if started:
		linechart[0].append(month)





	month += 1

model = LinearRegression()

model.fit(np.array(linechart[0]).reshape(-1, 1), linechart[1])

for extra in range(5):
	linechart[0].append(extra + month)

linechart.append([])

for m in linechart[0]:
	linechart[2].append(model.predict([[m]])[0])

workbook = xlsxwriter.Workbook(f'Report{student}{subject}.xlsx')
worksheet = workbook.add_worksheet(f"{student}{subject}")

chart = workbook.add_chart({'type': 'line'})

worksheet.write_row('A1', ['Month', 'Actual Page Number', 'Expected Page Number'])
worksheet.write_column('A2', linechart[0])
worksheet.write_column('B2', linechart[1])
worksheet.write_column('C2', linechart[2])

column = 'C'

chart.add_series({
	'name' : f'={student}{subject}!${column}$1',
	'categories' : f'={student}{subject}!$A$2:$A${len(linechart[0]) + 1}',
	'values': f'={student}{subject}!${column}$2:${column}${len(linechart[0]) + 1}',
	'line' : {
			'color' : 'red',
	  		'dash_type' : 'long_dash',
	  		'width' : 1
			}
	})

column = 'B'

chart.add_series({
	'name' : f'={student}{subject}!${column}$1',
	'categories' : f'={student}{subject}!$A$2:$A${len(linechart[0]) + 1}',
	'values': f'={student}{subject}!${column}$2:${column}${len(linechart[0]) + 1}',
	'line' : {
				'color' : 'blue'
			}
})



chart.set_title({'name' : f"{student}'s progress in {subject}"})
chart.set_x_axis({'name' : 'Month'})
chart.set_y_axis({'name' : 'Page Number'})


worksheet.insert_chart('F1', chart);

workbook.close()


