import xlrd 
import datetime
import mysql.connector 
from io import StringIO
import sys
from twilio.rest import Client

#Function that retrieves data from excel to a file for sql creation
def xlstosql():
	loc = ("kwa.xls") 
	wb = xlrd.open_workbook("kwa.xls") 
	sheet = wb.sheet_by_index(0) 
	f=open('sql.sql','w')
	#Appends each row as a ysql query to the file
	for i in range(1,137):
		a=sheet.cell_value(i,0)
		b=sheet.cell_value(i,1)
		c=sheet.cell_value(i,2)
		d=sheet.cell_value(i,3)
		k=c.split("-")
		h=d.split("-")
		str1=k[0]+k[1]
		str2=h[0]+h[1]
		if str1 and str2=="  ":
			n1='NULL'
			n2='NULL'
			f.write('INSERT INTO dates (name,phno,dob,dom) VALUES ("%s","%s","%s",%s");\n'%(a,b,str1+n1,str2+n2))
		else:
			f.write('INSERT INTO dates (name,phno,dob,dom) VALUES ("%s","%s","%s","%s");\n'%(a,b,str1,str2))
	f.close()

#Function that compares todays date with DB to get the index
#Compare the current Day with dom in DB	
def getDOB():
		mydb=mysql.connector.connect(host='localhost',database='kwa',user='root',password='Root@846');
		a=datetime.datetime.now()
		now="'"+a.strftime("%d")+a.strftime("%B")+"'"
		query="select name,phno from master where dob="+now
		c=mydb.cursor()
		c.execute(query)
		op=c.fetchall()
		if len(op)==0:
			print("No Birthdays found")
		else:
			for i in range(len(op)):
				name=op[i][0]
				phno=op[i][1]
				msg=wishM(name)
				twilio(msg,phno)

#Compare the current Day with dom in DB	
def getDOM():
		mydb=mysql.connector.connect(host='localhost',database='kwa',user='root',password='Root@846');
		a=datetime.datetime.now()
		now="'"+a.strftime("%d")+a.strftime("%B")+"'"
		query="select name,phno from master where dom="+now+";"
		c=mydb.cursor()
		c.execute(query)
		op=c.fetchall()
		if len(op)==0:
			print("No Marriage Days")
		else:	
			for i in range(len(op)):
				name=op[i][0]
				phno=op[i][1]
				msg=wishM(name)
				twilio(msg,phno)
		
#Retruns the phone number
def num(phno):
	return phno

#To get currrent day and month	
def getDay():
	a=datetime.datetime.now()
	now="'"+a.strftime("%B")+a.strftime("%d")+"'"
	return (now)

#Using twilio API for sending msg to phno
def twilio(msg,phno):
	account_sid = 'ACd5f6cfbf8a580795b6d39e541896cac8'
	auth_token = '8133fda061b31fe08e075cf11f618fee'
	client = Client(account_sid, auth_token)
	message = client.messages.create(
	to="+918688888933",from_='+12013315355',body=msg)
	print(message.sid)


#Sends the Marriage day message with the desired name
def wishM(name):
	msg="Dear "+name+",\nHappy Wedding Anniversary.\nBest Wishes from JN Bhaskar Secretary,KWA"
	return msg
#Sends the Birthday message with the desired name
def wishB(name):
	msg="Dear "+name+",\nHappy Birthday\nBest Wishes from JN Bhaskar Secretary,KWA"
	return msg
if __name__ == '__main__':
	getDOM()
	getDOB()


			



