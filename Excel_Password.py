##################################################################################
#Prerequisites:
# - win32com.client package need to install.
# - Give proper path for excel file.
# - 
#							
###################################################################################
########## Builtin Package
import sys

##########Thired Party
import win32com.client

openedDoc = win32com.client.Dispatch("Excel.Application")
filename= sys.argv[1]

password_file = open ( 'wordlist.lst', 'r' )
passwords = password_file.readlines()
password_file.close()

passwords = [item.rstrip('\n') for item in passwords]

# Result store Path
results = open('results.txt', 'w')

for password in passwords:
	print(password)
	try:
		wb = openedDoc.Workbooks.Open(filename, False, True, None, password)
		print("Successful! Please Enter your Password..[>] : "+password)
		results.write(password)
		results.close()
	except:
		print("Please varify your Credentials")
		pass