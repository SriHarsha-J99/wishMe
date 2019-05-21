import xlrd 
import datetime
import mysql.connector 
from io import StringIO
import sys
from twilio.rest import Client
import urllib
import urllib2
import time;

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
		query="select name,phno from master where dob="+getDay()
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
				print(msg) 	             
				#twilio(msg,phno)

#Compare the current Day with dom in DB	
def getDOM():
		mydb=mysql.connector.connect(host='localhost',database='kwa',user='root',password='Root@846');
		query="select name,phno from master where dom ="+getDay()
		c=mydb.cursor()
		c.execute(query)
		op=c.fetchall()
		if len(op)==0:
			print("No Marriage Days")
		else:	
			for i in range(len(op)):
				name=op[i][0]
				phno=op[i][1]
				print(name)
				print(phno)
				msg=wishM(name)
				msg91(msg,phno)
				time.sleep(12)
				
		
#Retruns the phone number
def num(phno):
	return phno

#To get currrent day and month	
def getDay():
	a=datetime.datetime.now()
	now="'"+a.strftime("%d")+a.strftime("%B")+"'"
	
	return now

def msg91(msg,phno):
	authkey = "277581Aezl0gNUW2j5ce2dfb7" # Your authentication key.
	mobiles = phno # Multiple mobiles numbers separated by comma.
	message = msg # Your message to send.
	sender = "KWAJNB" # Sender ID,While using route4 sender id should be 6 characters long.
	route = "4" # Define route
	values = {
          'authkey' : authkey,
          'mobiles' : mobiles,
          'message' : message,
          'sender' : sender,
          'route' : route
          }
  	values1 = {
          'authkey' : authkey,
          'mobiles' : '9248044456',
          'message' : message,
          'sender' : sender,
          'route' : route
          }      
	url = "http://api.msg91.com/api/sendhttp.php" # API URL
	postdata = urllib.urlencode(values) # URL encoding the data here.
	req = urllib2.Request(url, postdata)
	response = urllib2.urlopen(req)
	output = response.read() # Get Response
	print(output)
	time.sleep(12)
	postdata1 = urllib.urlencode(values1) # URL encoding the data here.
	req1 = urllib2.Request(url, postdata1)
	response1 = urllib2.urlopen(req1)
	output1 = response.read() # Get Response
	print(output1)

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


			



