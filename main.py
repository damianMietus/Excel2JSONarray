import Tkinter
import tkMessageBox
from Tkinter import *
import Tkinter, Tkconstants, tkFileDialog


root = Tk()

#Button commands
def directoryFunc():
	root.filename = tkFileDialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
	print (root.filename)
	dirEntry.delete(0,END)
	dirEntry.insert(0, root.filename)

def checkVar():

	
	errorStr = "Failed to execute for the following reasons:\n"
	errorFlag = False




	dirvar = dirv.get()
	print "Directory: "+dirvar

	# Directory error
	if not dirvar:
		errorStr= errorStr+"- No directory selected\n"
		errorFlag = True

	outpvar = outpv.get()
	print "Output filename: "+outpvar+".json"

	# Name.json error
	if not outpvar:
		errorStr= errorStr+"- No output file name given\n"
		errorFlag = True

	arvar = arv.get()
	print "Array name: "+arvar

	# Array name error
	if not arvar:
		errorStr= errorStr+"- No json array name given\n"
		errorFlag = True

	exvar = exv.get()
	print "Excel column: "+exvar

	mode = op1Select.get()
	print "Mode: "+mode

	# Column error
	if not exvar:
		errorStr= errorStr+"- No column given\n"
		errorFlag = True
	elif (exvar.isalpha() == False):
		errorStr= errorStr+"- Column may only be alphabetical\n"
		errorFlag = True	

	if (errorFlag == True):
		tkMessageBox.showinfo("Error", errorStr)
	else:
		print "Success"

root.geometry("500x175")
root.resizable(0, 0)
root.title("xlsx2JSONarray")
root.pack_propagate(0)

#leftFrame = Frame(root, width = 350, height = 300, bd = 1, relief=SUNKEN )
#leftFrame.pack(side =LEFT)
#leftFrame.pack_propagate(0)
rightFrame = Frame(root, width = 145, height = 150, bd = 2, relief=SUNKEN)
rightFrame.pack_propagate(0)
rightFrame.pack(side=RIGHT, anchor="n")

# Directory
dirLabel = Label(root, text="Choose file location:")
dirLabel.grid(row=0, column=0)

dirv = StringVar()
dirEntry = Entry(root, width=50,textvariable=dirv)
dirEntry.grid(row=3, column=0)

dirButton = Button(root, text="...", command=directoryFunc)
dirButton.grid(row=3, column = 1)

# Output name
outpLabel = Label(root, text="Output file name:")
outpLabel.grid(row=4, column = 0)

outpv = StringVar()
outpEntry = Entry(root, width=50, textvariable=outpv)
outpEntry.grid(row=5, column=0)

outpLabel = Label(root, text=".json")
outpLabel.grid(row=5, column = 1)

# Array name
arLabel = Label(root, text="Name of JSON array:")
arLabel.grid(row=6, column = 0)

arv = StringVar()
arEntry = Entry(root, width=50, textvariable=arv)
arEntry.grid(row=7, column=0)

# Excel Column
exLabel = Label(root, text="Excel column to read from:")
exLabel.grid(row=9, column = 0)

exv = StringVar()
exEntry = Entry(root, width=50, textvariable=exv)
exEntry.grid(row=10, column=0)





# Utilities frame

utilityLabel = Label(rightFrame, text="Text Format:",anchor="n")
utilityLabel.pack()

op1Vals = ["Default", "Lowercase", "Uppercase"]
op1Select = StringVar()
op1Select.set(op1Vals[0])

op1 = OptionMenu(rightFrame, op1Select, *op1Vals)
op1.pack()



# The Execute Button

excLabel = Label(rightFrame, text="")
excLabel.pack()
executeButton = Button(rightFrame, text="Execute", bd = 3, width=10, height=3,pady = 3, command = checkVar)
executeButton.pack(anchor="s")

root.mainloop()


