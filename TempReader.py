#indention is 4 spaces, not tabs
import urllib.request
import time
import xlsxwriter
import tkinter
import multiprocessing

#Create New Test Window
def StartTestClick(pipeArray):
    NewTestInfo = [' ', ' ', ' ', ' ']  # [ProdName, Lot#, Test#/Letter, Probe#]
    ok = tkinter.IntVar()
    radBut = tkinter.IntVar()

    newTest = tkinter.Toplevel()
    newTest.geometry('350x125')
    newTest.title("New Test")
    lbl1 = tkinter.Label(newTest, text="Product Name")
    lbl2 = tkinter.Label(newTest, text="Lot #")
    lbl3 = tkinter.Label(newTest, text="Test #/Letter")
    lbl4 = tkinter.Label(newTest, text="Probe Number")
    lbl1.grid(column=0, row=0, sticky='E')
    lbl2.grid(column=0, row=1, sticky='E')
    lbl3.grid(column=0, row=2, sticky='E')
    lbl4.grid(column=0, row=3, sticky='E')

    txt1 = tkinter.Entry(newTest, width=15)
    txt2 = tkinter.Entry(newTest, width=15)
    txt3 = tkinter.Entry(newTest, width=5)
    txt1.grid(column=1, row=0)
    txt2.grid(column=1, row=1)
    txt3.grid(column=1, row=2, sticky='W')

    rad1 = tkinter.Radiobutton(newTest, text='1', value=1, variable=radBut)
    rad2 = tkinter.Radiobutton(newTest, text='2', value=2, variable=radBut)
    rad3 = tkinter.Radiobutton(newTest, text='3', value=3, variable=radBut)
    rad4 = tkinter.Radiobutton(newTest, text='4', value=4, variable=radBut)
    rad5 = tkinter.Radiobutton(newTest, text='5', value=5, variable=radBut)
    rad6 = tkinter.Radiobutton(newTest, text='6', value=6, variable=radBut)
    rad1.grid(column=1, row=3, sticky='E')
    rad2.grid(column=2, row=3)
    rad3.grid(column=3, row=3)
    rad4.grid(column=4, row=3)
    rad5.grid(column=5, row=3)
    rad6.grid(column=6, row=3)

    rad1.select()
    txt1.focus()
	
    Space1 = tkinter.Frame(newTest)
    Space1.grid(column=0, row=4)

    def okClick():
        ok.set(1)
    def cancelClick():
        ok.set(0)

    but1 = tkinter.Button(newTest, text='OK', width=10, command=okClick)
    but1.place(height=26, width=90, relx=0.15, rely=0.71)
    but2 = tkinter.Button(newTest, text='Cancel', width=10, command=cancelClick)
    but2.place(height=26, width=90, relx=0.50, rely=0.71)

    newTest.focus()
    txt1.focus()

 #print(ok.get())

    but1.wait_variable(ok)

    #print(ok.get())

    NewTestInfo[0] = txt1.get()
    NewTestInfo[1] = txt2.get()
    NewTestInfo[2] = txt3.get()
    NewTestInfo[3] = radBut.get()

    if(ok.get() == 1):
        p = multiprocessing.Process(target=TestStart, args=(NewTestInfo, pipeArray[NewTestInfo[3]-1]))    # -1 since NewTestInfo is an array that starts at [0] and probe numbers start at 1
        p.start()
    if(ok.get() == 0):
        print("User Canceled")

    newTest.destroy()

def TestStart(NewTestInfo, child_conn):
    # Starting point for recording test data
    fileName = fileNamer(NewTestInfo[0], NewTestInfo[1], NewTestInfo[2])
    # initializes the excel workbook and sheet
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold': True})

    worksheet.write('A1', 'TIME', bold)
    worksheet.write('B1', 'TEMP', bold)
    worksheet.write('D1', 'PRODUCT', bold)
    worksheet.write('E1', 'LOT', bold)
    worksheet.write('D2', NewTestInfo[0])
    worksheet.write('E2', NewTestInfo[1])
    worksheet.write('D4', 'FINAL WT', bold)
    worksheet.write('D5', 'FINAL RT', bold)
    worksheet.write('D6', 'FINAL EXO', bold)
    # After initializing the excel sheet, we start recording
    finalTimes = [0.0, 0.0, 0]
    child_conn.send('testing')
    #child_conn.close()
    finalTimes = timeKeeper(NewTestInfo[3], worksheet, child_conn)

    for i in (0, 1):
        finalTimes[i] = timeStandardizer(finalTimes[i])
    # writes final times/temp on excel sheet
    worksheet.write('E4', finalTimes[0], bold)
    worksheet.write('E5', finalTimes[1], bold)
    worksheet.write('E6', finalTimes[2], bold)

    workbook.close()

    child_conn.send('ready')
    child_conn.close()

    return

def fileNamer(prodName, prodLot, testNum):
    # assembles the string responsible for excel file name and location.
    fileName = 'C://USERS/USER/Desktop/ProgTest/' + prodName + '_' + prodLot + '_' + testNum + '.xlsx'
    return fileName

def timeKeeper(pNumber, worksheet, child_conn):
    # Counts Time for infoGrabber and Final Test Times
    flag = '1'
    count = 1
    tempMax = int('0')
    finalTimes = [0.0, 0.0, 0]  # [WT, RT, EXO]
    startTime = time.time()

    while (flag == '1'):
        time.sleep(.75)  # stops the function for 0.75 seconds to cut back how often we are requesting updates from the Yokogawa
        curTemp = infoGrabber(pNumber)
        count = eXcelWriter(time.time() - startTime, curTemp, count, worksheet)  # writes current sec count since start and current temp, function returns count + 1 to ensure next excel cell is used
        if (curTemp > tempMax + 100):
            child_conn.send('error')
            #child_conn.close()
        if (curTemp > tempMax):  # updates max temp and time in an array
            tempMax = curTemp
            finalTimes[1] = time.time() - startTime
            finalTimes[2] = tempMax
        if (curTemp >= 105 and finalTimes[0] == 0.0):  # records WT
            finalTimes[0] = time.time() - startTime
            child_conn.send('hot')
            #child_conn.close()
            print("105F TIMES MET/RECORDED!")  # Only used during testing
        if ((tempMax - 10) > curTemp and curTemp < tempMax):  # breaks the while loop to end the function
            flag = '0'
    return (finalTimes)

def infoGrabber(pNumber):
    # Requests current temperature information from Yokogawa DX100P and returns the requested probe temp
    with urllib.request.urlopen('http://192.168.1.137/cgi-bin/ope/allch.cgi') as f:
        raw_data = f.read()  # converts source information from yokogawa into a string
    if (pNumber == 1):  # based on probe number, cuts all information except requested probe temp from the string
        raw_data = raw_data[2269:-4035]
        raw_data = int(raw_data.decode('utf-8'))
        return raw_data
    elif (pNumber == 2):
        raw_data = raw_data[3056:-3248]
        raw_data = int(raw_data.decode('utf-8'))
        return raw_data
    elif (pNumber == 3):
        raw_data = raw_data[3842:-2461]
        raw_data = int(raw_data.decode('utf-8'))
        return raw_data
    elif (pNumber == 4):
        raw_data = raw_data[4629:-1674]
        raw_data = int(raw_data.decode('utf-8'))
        return raw_data
    elif (pNumber == 5):
        raw_data = raw_data[5416:-887]
        raw_data = int(raw_data.decode('utf-8'))
        return raw_data
    elif (pNumber == 6):
        raw_data = raw_data[6203:-100]
        raw_data = int(raw_data.decode('utf-8'))
        return raw_data

def eXcelWriter(curTime, curTemp, count, worksheet):
    # Writes read data from Yokogawa to excel sheet
    worksheet.write(count, 0, curTime)
    worksheet.write(count, 1, curTemp)
    print("Writing to Excel File")  # Only used in testing
    return count + 1

def timeStandardizer(time):
    # Converts from decimal minutes to min:sec
    x = 0.0
    time = (time - (time % 1))
    x = time / 60
    x = (x - (x % 1))
    if (int(time - (x * 60))) < 10:
        time = str(int(x)) + ':' + '0' + str(int(time - (x * 60)))	
    else:
        time = str(int(x)) + ':' + str(int(time - (x * 60)))
    return time
