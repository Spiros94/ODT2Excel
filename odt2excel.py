import Tkinter, tkFileDialog
import xlsxwriter
import os
import sys

targetFileName = 'graph'
targetFileExt = '.xlsx' # Currently no support for other extensions
targetOutDir = './output/'
ODTfileIn = 'graph.odt'

def CheckInputFile(ODTfile):
    if not os.path.isfile(ODTfile):
        print "[!] No file named " + ODTfile + " found. Exit"
        END = raw_input('..Press Enter key to exit..')
        raise SystemExit

class XlsxFileManage:

    def __init__(self,Filename,Extension,OutputDir):
        self.filename = Filename
        self.extension = Extension
        self.outputdir = OutputDir
        self.CheckOutputPathExistance()
        self.CheckOutputFileExistance()

    def FullPathedFile(self):
        # Returns the file with the path
        return './' + self.outputdir + '/' + self.filename + self.extension

    def MakeBookNSheet(self):
        # Makes the book and the sheet using MakeWorkBook() and MakeWorkSheet()
        self.MakeWorkBook()
        self.MakeWorkSheet()
        
    def MakeWorkBook(self):
        # Makes a workbook (NOT a sheet for it)
        self.workbookfile = xlsxwriter.Workbook(self.FullPathedFile())
        
    def MakeWorkSheet(self):
        # Makes a sheet for the workbook
        self.worksheet = self.workbookfile.add_worksheet()

    def WorkBookWrite(self,row,col,item):
        # Writes a shell in the sheet
        self.worksheet.write(row,col,item)

    def WorkBookClose(self):
        # Saves/Closes the workbook
        self.workbookfile.close()

    def CheckOutputFileExistance(self):
        # Checks if the output file already exists. If so it creates another one with a different name
        checkfile = self.outputdir + self.filename + self.extension  # Concat output directory - filename.extension
        i = 1 # Just an iterator
        while os.path.isfile(checkfile): # If output file exists add a number at then end of the filename and try again
            self.filename += str(i)
            checkfile = self.outputdir + self.filename + self.extension
            i += 1

    def CheckOutputPathExistance(self):
        # Check if 'output' folder exists. If not create one
        if not os.path.exists(self.outputdir):
            os.makedirs(self.outputdir)

def Row4Split(row4):
    # Split and concat the fourth row of the file
    row4new = list() # List for the splitted row
    row4 = row4.replace(row4[:12],'') # Remove the first 12 characters
    row4 = row4.split(' ')
    row4 = filter(lambda name: name.strip(), row4) # Remove spaces from list
    colbuffer = ''
    for column in row4:
        if column[0] != '{' and colbuffer == '':
            row4new.append(column)
        else:
            if '}' in column:
                colbuffer = colbuffer + column
                row4new.append(colbuffer)
                colbuffer = ''
            else:
                colbuffer = colbuffer + column + '_'
    return row4new

def Row5Split(row5):
    row5 = row5.replace(row5[:8],'') # Remove the first 8 characters
    row5 = row5.split(' ')
    row5 = filter(lambda name: name.strip(), row5) # Remove spaces from list
    return row5

def NumberRowSplit(nrow):
    nrow = nrow
    nrow = nrow.split(' ')
    nrow = filter(lambda name: name.strip(), nrow)
    return nrow
'''
# Console based question
def AskForInput(ODTfile):
    q = raw_input('[?] Enter file name(default graph.odt): ')
    if q != '':
        return q
    else:
        return ODTfile
'''

def main():
    root = Tkinter.Tk()
    root.withdraw()
    ODTfile = tkFileDialog.askopenfilename()
    print '[+] Open File'
    
    ''' # Console based question
    ODTfile = AskForInput(ODTfileIn)
    '''
    
    print "[+] Input File: " + ODTfile
    CheckInputFile(ODTfile) # Check if graph.odt exists on the same folder
    Infile = open(ODTfile,'r') # Open odt file
    print '[+] File is open. Trying to count lines'
    num_lines = sum(1 for line in open(ODTfile)) # Count how many lines are in the file
    OutputFile = XlsxFileManage(targetFileName,targetFileExt,targetOutDir)
    OutputFile.MakeBookNSheet()
    
    # Write the first 3 rows to the xlsx file
    for i in range(0, 3):
        OutputFile.WorkBookWrite(i,0,Infile.readline())
        
    print '[+] Columns Extraction'
    
    ColsNames = Row4Split(Infile.readline())
    for i in range(0,len(ColsNames)):
        OutputFile.WorkBookWrite(3,i,ColsNames[i])

    print '[+] Units Extraction'
    Units = Row5Split(Infile.readline())
    for i in range(0,len(Units)):
        OutputFile.WorkBookWrite(4,i,Units[i])
    print '[+] Values Extraction'
    for i in xrange(5, num_lines): # Table End line will also be an output here, splitted
        nthrow = NumberRowSplit(Infile.readline())
        for y in range(0, len(nthrow)):
            OutputFile.WorkBookWrite(i,y,nthrow[y])

    OutputFile.WorkBookClose()
    Infile.close()
    print "[+] End"
    END = raw_input('..Press Enter key to exit..')

if __name__ == "__main__":
       main()
