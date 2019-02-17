###############################################################
## This module contains all the generic/general modules that
## are called by multiple programs.
###############################################################
from datetime import date, datetime
import pypyodbc
import LoginFrameIntegratedEmail
import os
import bz2
from datetime import datetime
from time import sleep

def writeMultiXlsx(filename,sheetNameArray,dataArray,headerColour):

    #Create XLS
    workbook = xlsxwriter.Workbook('' + filename + '.xlsx', {'constant_memory': True})
    format = workbook.add_format({'bold': True, 'font_color': ''+headerColour+''})

    # Create Dynamic callable Var Sheets for each input split
    for x, name in enumerate(sheetNameArray):
        locals()['sheet{}'.format(x)] = workbook.add_worksheet(sheetNameArray[x])
        locals()['sheet{}'.format(x)].set_default_row(15)
        locals()['sheet{}'.format(x)].set_row(0, 15, format)
        locals()['sheet{}'.format(x)].freeze_panes(1, 0)

        # Write Data Headings for each sheet
        for y, heading in enumerate(dataArray[0]):
            locals()['sheet{}'.format(x)].write(0,y,heading)

    dataArray.__delitem__(0)

    arrElement = 0
    col=0
    r=1
    sheetLoops = dataArray.__len__()/sheetNameArray.__len__()
    remainder = dataArray.__len__()% sheetNameArray.__len__()
    print ("remainder"+str(remainder))

    #place data in sheets
    for loop in range (0,int(sheetLoops)):
        for currentSheet in range (0,int(sheetNameArray.__len__())):
            for c, arr in enumerate(dataArray[col]):
                locals()['sheet{}'.format(currentSheet)].write(r, c, dataArray[arrElement][c])
            arrElement+=1
        r+=1

    for currentSheet in range (0,int(remainder)):
        for c, arr in enumerate(dataArray[col]):
            locals()['sheet{}'.format(currentSheet)].write(r, c, dataArray[arrElement][c])
        arrElement+=1
    r+=1
    
def none_to_space(variable):
    """This function converts a variable that is of None-type to an empty string.
    """
    if variable is None:
        return ''
    else:
        return variable


def populatefile(fieldsbefore, fieldsafter, previous_records, separator):
    """This function merges and populates a file.

    Input Keyword Arguments:
    fieldsbefore: Fields that must appear at the start of every line. This can
                  be an empty string or a single line string.
    fieldsafter: Fields that must appear at the end of every line. This can be
                 an empty string or a single line string.
    previous_records: Records that must appear on the new file, prefixed with the
                      fieldsbefore and suffixed with the fieldsafter.  This can be
                      a multi-line string, an empty string or a single line string.
    separator: The field separator in the file e.g. ','.

    Values Returned:
    1.  newfile_records: New file's records

    Example:
    fieldsbefore = 'A, B'
    fieldsafter  = 'C'
    previous_records = '1,2\n
                        3,4\n'
    separator = ','
    newfile_records = 'A,B,1,2,C\n
                       A,B,3,4,C\n'
    """
    newfile_records = ''
    fieldsbefore = none_to_space(fieldsbefore).strip()
    fieldsafter = none_to_space(fieldsafter).strip()
    previous_records = none_to_space(previous_records)
    if fieldsbefore != '' and fieldsbefore[-1] != separator:
        fieldsbefore += separator
    if fieldsafter != '' and fieldsafter[0] != separator:
        fieldsafter = separator + fieldsafter
    if previous_records.strip() == '':
        newfile_records = fieldsbefore + fieldsafter[1:] + '\n'
    else:
        lines = previous_records.split('\n')
        for line in lines:
            if line != '':
                ##                if line[-1] == separator:
                ##                    line = line[0:-1]
                newfile_records += fieldsbefore + line + fieldsafter + '\n'

    return newfile_records

def populateLibraryDictionary(libraryFileName):
    """This function populates a dictionary that contains the File
    and the Library in which it is located.

    Input Keyword Arguments:
    libraryFileName: The text file that contains the list of Files and their respective Libraries.
                     The File and its Library must be separated by a comma (,).
                     If a File is repeated, then the last Library associated with it is returned.

    Values Returned:
    1.  A dictionary containing the File as the key and its associated library as the data
    """
    libraryDictionary = {}
    with open(libraryFileName) as inputFile:
        for line in inputFile:
            if line.strip() != '':
                lineSplit = line.split(',')
                libraryDictionary[lineSplit[0].strip()] = lineSplit[1].strip()
    return libraryDictionary

def financialDateCalc(addMonths=0,dateOverride='CCYYMMDD or DD-MM-CCYY',returnInteger=False, invertDateFormat=False, setDayOfMonth=1):
    """By default this function returns the current findate as atetime and will add or subtract
    the months you provided as an input, the function also has some advanced date manipulation functions:
    :param addMonths: int - Allows the user the change the returned date by adding or subtracting months
    :param dateOverride: String - Allows the user to pass in multiple string formats of dates which overrides the current date used
    :param returnInteger: bool - Changes the output return between date or int format
    :param invertDateFormat: bool - swaps the output dates day and year positions
    :param setDayOfMonth: int - Allows the user to override the output dates day
    :returns date / int of date
    """
    m=0
    y=0
    c=0

    if dateOverride=='CCYYMMDD or DD-MM-CCYY' or dateOverride== None or dateOverride== False:
        d=date.today()
    elif dateOverride.__len__()==8:
        try:
            d=datetime(year=int(dateOverride[0:4]),month=int(dateOverride[4:6]),day=int(dateOverride[6:8]))
            setDayOfMonth=int(dateOverride[6:8])
        except ValueError:
            d=datetime(day=int(dateOverride[0:2]),month=int(dateOverride[2:4]),year=int(dateOverride[4:8]))
            setDayOfMonth=int(dateOverride[0:2])
    elif dateOverride.__len__()==10:
        try:
            d=datetime(year=int(dateOverride[6:10]),month=int(dateOverride[3:5]),day=int(dateOverride[0:2]))
            setDayOfMonth=int(dateOverride[0:2])
        except ValueError:
            d=datetime(year=int(dateOverride[0:4]),month=int(dateOverride[5:7]),day=int(dateOverride[8:10]))
            setDayOfMonth=int(dateOverride[8:10])
    else:
        raise TypeError('Error with input')

    while (addMonths > 12):
        addMonths-=12
        c+=1

    if (d.month+addMonths) > 12:
        m+=(d.month+addMonths - 12)
        y+=(d.year+1)
    else:
        m += (d.month+addMonths)
        y += d.year

    i=0
    while True:
        try:
            dateR=date(y+c, m, setDayOfMonth-i)
        except ValueError:
            i+=1
        else:
            break
    if returnInteger==False:
        if invertDateFormat==False:
            return dateR
        else:
            return dateR.strftime('%d-%m-%Y')
    else:
        if invertDateFormat==False:
            return int(str(dateR.strftime('%Y%m%d')))
        else:
            return int(str(dateR.strftime('%d%m%Y')))

def tkinterCenter(self, win):
    '''
    This function centre's TK frame according to the size of the users screen resolution
    :param win: Send in the root TK frame object
    :return: None
    '''
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()


def encyptUsernamePassword(system,username,password):
    '''
    This simple function uses a basic compression library to encrypt system login, usernames and passwords.
    :param system
    :param username
    :param password
    :return: None - function creates the Session.dat file
    '''
    computer_name = os.environ['COMPUTERNAME']
    compressed = bz2.compress(bytes((system+'\0'+username + '\0' + password + '\0' + computer_name).encode('utf-8')),9)
    with open("Session.dat", "wb") as f:
        f.write(compressed)
        f.close()

def decryptUsernamePassword():
    '''
    This simple function uses a basic compression library to decrypt a file created with encyptUsernamePassword method
    '''
    try:
        with open("Session.dat", "rb") as f:
            d_data = bz2.decompress(f.read()).decode('latin')
            f.close()
        decypher=(d_data.split('\0'))

        if decypher[3] != os.environ['COMPUTERNAME']:
            #print("Error Detected: Cypher Key different from provided Key, Deleted stored Username and Password")
            with open("Session.dat", "wb") as f:
                f.write(bz2.compress(bytes((" ").encode('utf-8')),9))
                f.close()
            return True
    except:
        return False
    return decypher[0],decypher[1],decypher[2]

def createCSVfromList(saveFileName,list):
    '''
    This function creates and saves a CSV file from a basic list
    :param saveFileName: string - name of the save filename
    :param list: [] the comma delimetered list
    '''
    count=0
    with open(str(saveFileName)+'.csv', 'w') as myfile:
        for item in list:
            count+=1
            if count ==list.__len__():
                myfile.write(str(item))
            else:
                myfile.write(str(item) +',')
    myfile.close()

def createListFromCSV(saveFileName, orientation=0):
    '''
    This function loads a CSV file and converts it into a list
    :param saveFileName: string - name of the save filename to load
    :return list[]
    '''
    inputList =[]
    with open(saveFileName+'.csv', 'r') as f:
        if orientation == 0:
            for line in f:
                stippedLine = line.split(',')
                inputList += stippedLine

        elif orientation == 1:
            for line in f:
                inputList.append(line.strip("\n"))
        else:
            raise TypeError('The provided orientation does not exist')

    f.close()
    return inputList

def bool(input):
    '''
    This function parses(converts) various variable strings into the correct boolean
    :param input:string - true or 1 or yes or t or false or 0 or no or f
    :return bool: True or False
    '''
    input=str(input.lower())
    if input == 'true' or input == '1' or input == 'yes' or input == 't':
        output = True
    elif input == 'false' or input == '0' or input == 'no' or input == 'f':
        output = False
    else:
        raise TypeError('Cannot convert the provided input into bool')
        return
    return output

def replaceCharInString(string,replaceThis,withThat):
    '''
    This function takes a string and replaces all occurrences of an input char with another input char
    :param string: string - Input
    :param replaceThis: char - string - Char to be replaced
    :param withThat: char - string - Char that is inserted in place of replaced
    :return newString
    '''
    re = set(replaceThis)
    newChar = set(withThat)
    newString=''
    for c in string:
        if c in re:
            newString+=newString.join(newChar)
        else:
            newString+=newString.join(c)
    return newString

def removeCharsFromString(string, chars):
    '''
    This function takes a removes all occurrences of the provided characters in a string
    :param string: string - Input
    :param chars: all characters to be removed i.e. 'i*a'
    :return newString
    '''
    sc = set(chars)
    return ''.join([c for c in string if c not in sc])

def wait_start(runTime,waitInterval=10):
    '''
    This function takes a string version of a datetime and forces the program till wait until the system time is equal to the provided date and time
    :param runTime: string - in datetime format "%d-%m-%Y %H:%M:%S"
    :param waitInterval: int - changes the frequncy in seconds that the method checks current time vs runtime
    '''
    startTime = datetime.strptime(runTime, "%d-%m-%Y %H:%M:%S")
    print("Waiting to start @ " + str(startTime))
    while startTime > datetime.today():
        sleep(waitInterval)
    return

def binarySearch(sortedList, searchItem):
    '''
    This function takes prtforms a search
    :param sortedList: sorted list
    :param searchItem: item to search for
    :returns Bool: Ture if item is found else false if item is not found
    :returns int: Index of the found item
    '''
    first = 0
    last = len(sortedList)-1
    found = False
    midpoint=-1
    #searchItem

    while first<=last and not found:
        midpoint = (first + last)//2
        if str(sortedList[midpoint]) == str(searchItem):
            found = True
        else:
            if str(searchItem) < str(sortedList[midpoint]):
                last = midpoint-1
            else:
                first = midpoint+1

    return found,midpoint