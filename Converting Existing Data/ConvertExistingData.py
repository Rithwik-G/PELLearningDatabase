import mysql.connector
import pandas as pd
import random

db = mysql.connector.connect()

db = mysql.connector.connect(
	host = "localhost",
	user = "admin",
	passwd = "1234",
	database = "StudentData"
	)

mycursor = db.cursor()

i = 0

while True:
	i += 1

	try:
		mycursor.execute(f"DROP TABLE StudentDataMonth{i}")
	except:
		pass
	
	try:
		data = pd.read_excel("StudentData.xlsx", index_col = "No.", sheet_name = f"Month{i}")
	except:
		print(f"Gathered Data until Month {i - 1}")
		break;

	
	try:
		
		mycursor.execute(f"CREATE TABLE StudentDataMonth{i} (No smallint PRIMARY KEY AUTO_INCREMENT, Subject VARCHAR(1), FirstName VARCHAR(50), Tel VARCHAR(50), Email VARCHAR(50), DOE VARCHAR(15), PEL_Wks_Level VARCHAR(3), PEL_Wks_No VARCHAR(10), Notes VARCHAR(100), CODE VARCHAR(1), Date VARCHAR(15))")
	except mysql.connector.errors.ProgrammingError:
		pass

	for ind in range(1, len(data) + 1):

		mycursor.execute(f"INSERT INTO StudentDataMonth{i} (Subject, FirstName, Tel , Email , DOE, PEL_Wks_Level, PEL_Wks_No, Notes, CODE, Date) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", list(data.loc[ind]))

db.commit()




