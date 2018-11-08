#import sys
#import string
#import time


import Tkinter
from Tkinter import *
import Tkinter, Tkconstants, tkFileDialog

root = Tk()

root.geometry("400x200")
root.resizable(0, 0)
root.title("xlsx2JSONarray")
root.pack_propagate(0)

leftFrame = Frame(root)
leftFrame.pack(side =LEFT)
rightFrame = Frame(root, width = 150, height = 100, bd = 2, relief=SUNKEN )
rightFrame.pack_propagate(0)
rightFrame.pack(side=RIGHT, anchor="n")

'''
rightTopFrame = Frame(root)
rightTopFrame = Frame(root, width = 100, height = 100, bd = 2, relief=SUNKEN )
rightTopFrame.pack(rightFrame)
'''

button1 = Button(leftFrame, text="Button 1", fg="red")
button2 = Button(rightFrame, text="Button 2", fg="purple")

button1.pack()
button2.pack()



dirButton = Button(leftFrame, text="...")
dirButton.pack()





op1Vals = ["Default", "Set to Lowercase", "Set to Uppercase"]
op1Select = StringVar()
op1Select.set(op1Vals[0])

op1 = OptionMenu(rightFrame, op1Select, *op1Vals)
op1.pack()

#root.filename = tkFileDialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
#print (root.filename)


root.mainloop()


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
