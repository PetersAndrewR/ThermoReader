# Yokogawa DX100P Temp Reader and Recorder for PC
# 14MAY2018 by Andrew Peters
#indention is tabs
import TempReader
import multiprocessing
import tkinter
from tkinter import messagebox
import os

def exampleOpener():
	#Opens the Example output xlsx for the user to see
	file = "C://Users/USER/Desktop/ProgTest/Example_Output.xlsx"
	os.startfile(file)
	return

def readmeOpener():
	#Opens the Readme for User Instruction and Information
	file = "C://Users/USER/Desktop/ProgTest/readme.txt"
	os.startfile(file)
	return
	
def statusDecoder(status):
	#Determines what image to use when updating main window	
	global image0, image1, image2, image3
	if status == 'ready':
		return image0
	elif status == 'testing':
		return image1
	elif status == 'hot':
		return image2
	elif status == 'error':
		return image3
	else:
		print("Problem with images")
		return
		
def refresh():
	#Updates the images on the main Window that show Probe Status
	global parent_conn1, parent_conn2, parent_conn3, parent_conn4, parent_conn5, parent_conn6
	if parent_conn1.poll():
		status = parent_conn1.recv()
		label0.config(image=statusDecoder(status))
	if parent_conn2.poll():
		status = parent_conn2.recv()
		label1.config(image=statusDecoder(status))
	if parent_conn3.poll():
		status = parent_conn3.recv()
		label2.config(image=statusDecoder(status))
	if parent_conn4.poll():
		status = parent_conn4.recv()
		label3.config(image=statusDecoder(status))
	if parent_conn5.poll():
		status = parent_conn5.recv()
		label4.config(image=statusDecoder(status))
	if parent_conn6.poll():
		status = parent_conn6.recv()
		label5.config(image=statusDecoder(status))
	print("REFRESHING!")
	root.after(1000, refresh)

if __name__ == '__main__':    #Prevents new process' from opening a new main window when a test is started
	
	#Duplex Pipe assignment, the parents stay with the main window, the child_conn is passed to the test process.
	parent_conn1, child_conn1 = multiprocessing.Pipe()
	parent_conn2, child_conn2 = multiprocessing.Pipe()
	parent_conn3, child_conn3 = multiprocessing.Pipe()
	parent_conn4, child_conn4 = multiprocessing.Pipe()
	parent_conn5, child_conn5 = multiprocessing.Pipe()
	parent_conn6, child_conn6 = multiprocessing.Pipe()

	pipeArrayC = [child_conn1, child_conn2, child_conn3, child_conn4, child_conn5, child_conn6]
	pipeArrayP = [parent_conn1, parent_conn2, parent_conn3, parent_conn4, parent_conn5, parent_conn6]

#Main Window
	root = tkinter.Tk()
	root.geometry('310x425')
	root.title("ThermoReader9001")
	root.columnconfigure(0, minsize=150)

	btn1 = tkinter.Button(root, text="Start Test", command=lambda: TempReader.StartTestClick(pipeArrayC))  #The lambda expression allows me to pass variables to the function when the button is clicked.
	btn2 = tkinter.Button(root, text="Example Output", command=exampleOpener)
	btn3 = tkinter.Button(root, text="Open ReadMe", command=readmeOpener) 

	btn1.grid(column=0, row=1)
	btn2.grid(column=0, row=3)
	btn3.grid(column=0, row=2)
	
	Space1 = tkinter.Frame(root)
	Space3 = tkinter.Frame(root)
	Space5 = tkinter.Frame(root)
	Space1.grid(column=0, row=0)
	Space3.grid(column=0, row=2)
	Space5.grid(column=0, row=4)

	image0 = tkinter.PhotoImage(file="C://Users/USER/Documents/ready.png")
	image1 = tkinter.PhotoImage(file="C://Users/USER/Documents/test_run.png")
	image2 = tkinter.PhotoImage(file="C://Users/USER/Documents/temp_met.png")
	image3 = tkinter.PhotoImage(file="C://Users/USER/Documents/probe_er.png")

	label0 = tkinter.Label(image=image0)
	label1 = tkinter.Label(image=image0)
	label2 = tkinter.Label(image=image0)
	label3 = tkinter.Label(image=image0)
	label4 = tkinter.Label(image=image0)
	label5 = tkinter.Label(image=image0)

	label0.grid(column=1, row=0, pady=5)
	label1.grid(column=1, row=1)
	label2.grid(column=1, row=2)
	label3.grid(column=1, row=3)
	label4.grid(column=1, row=4)
	label5.grid(column=1, row=5)
	'''
	messagebox.showinfo("Quick Guide",
                    "TempReader will record the temperature for a specified probe every second as well as determining the WT, CT, and EXO.  "
                    "Temp Reader will then save the resulting information to an appropriately named excel file.\n\nAn example is packaged with this program.  "
                    "To view press the Example button on the main window.\n\nThings to Consider:\n\nIf you close the main menu, TempReader will cease "
                    "all recording.  Leave the Main Window open until all tests are complete.\n\nThe program decides when the test is over and to stop "
                    "recording when the temperature drops 10 degrees.  If you cancel a test early and a drop of 10 degrees or more does not occur TempReader"
                    "will continue to record.  Simple remedy is to dip the probe in cold water.")
	'''
	refresh()
	
	root.mainloop()
