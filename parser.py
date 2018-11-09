#import sys
#import string
#import time

'''
# Check the imports
try:
	from openpyxl import load_workbook
	from openpyxl import Workbook
except:
	print ("Could not import openpyxl.")
	time.sleep(5)
	sys.exit(1)

try:
	from tqdm import tqdm
except:
	print ("Could not import tqdm.")
	time.sleep(5)
	sys.exit(1)
	
# Get the directory of the excel file
while (True):
	filename = raw_input("Please input the name of the Excel file: ")
	try:
		wb = load_workbook(filename, data_only = True)
		ws = wb.active
		print ("Successfully opened the workbook")
		break
	except:
		print ("error opening the file")

print("Its done")
time.sleep(5)
sys.exit(1)
'''