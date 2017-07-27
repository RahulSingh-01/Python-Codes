import PyPDF2 as PDF
import getpass
import os
import argparse
import logging
import time
import shutil
import zipfile
import re
import sys
import subprocess
from datetime import datetime
import pandas as pd

#Name of the windows application for ghostscript
GHOSTSCRIPTCMD = 'gswin32c'

"""
Parameter to run the script
python pdfsplitter.py --sourcefiledirectory "C:\\Temp\PDF\Source" --targetfiledirectory "C:\\Temp\PDF\Target"

"""
def create_logger():
    logging.basicConfig(filename='pdfsplitter.log',level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

def log(msg):
    logging.info(msg)
    print('[' + time.strftime('%Y-%m-%d %H:%M:%S') + '] ' + msg)

create_logger()

log('Process initiated by user {}'.format(getpass.getuser()))

try :

    log('Processing the PDF Split .........')
    

    def makeArchive(fileList, archive, root):
        a = zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED)
        for f in fileList:
            print("archiving file %s" % (f))
            a.write(f, os.path.relpath(f, root))
            os.remove(f)
        a.close()

    def compresspdf(pdfout,pdfin):
        try:
            """
            change the pdfsetting to ebook to reduce the size by compressing white spaces
            call Ghostscipt via commandline i.e. Subprocess command
            """
            arglist = [GHOSTSCRIPTCMD,"-sDEVICE=pdfwrite","-dNOPAUSE","-dBATCH","-dPDFSETTINGS=/ebook","-sOutputFile=%s" %pdfout,"%s" %pdfin] 
            sp = subprocess.Popen(args=arglist,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
            stdout, stderr = sp.communicate()
            os.remove(pdfin)
            os.rename(pdfout,pdfin)
        except OSError:
            sys.exit("Error executing Ghostscript ('%s'). Is it in your PATH?" %GHOSTSCRIPTCMD)

    

    def getstaffno_page(pdfread,staffstr = 'STAFF(\w{7})',paynostr = 'PAYNO(\d*)',paydatestr = 'PAYDATE(\w{9})' ):
        
        """
        staffstr - String Serach Parameter for Staff Number
        paynostr - String Search Parameter for Pay Run Number
        paydatestr - String Search Parameter for Pay Date
        This function performs string serach - REGEX for Staff number , Pay Number and Paydate on each page of PDF and 
        creates a dictionary object with list of pages for each key group of Staff,Pay run and Paydate
        """
        d = {}
        numpages = pdfread.numPages
        for i in range(numpages):
            pageobj = pdfread.getPage(i)
            a = pageobj.extractText()
            a = re.sub('[^a-zA-Z\d]','',a) #Replace any non digit or non aplhabet character to none
            stfno = re.search(staffstr,a,re.I)
            payno = re.search(paynostr,a,re.I)
            paydt = re.search(paydatestr,a,re.I)
            key = stfno.group(1)+'_'+payno.group(1)+'_'+paydt.group(1)
            if not key in d.keys() :
                d[key] = [i]
            else:
                d[key].append(i)
        return(d) 


    def writeemployeepaylog(f,dict):
        """ Function to write the Employee Number and pages they are found in for logging """
        f = open(flog,'w')
        for i in dict:
            val = [x+1 for x in dict[i]]
            data = str(i)+ ":" + str(val)
            f.write(data)
            f.write("\n")
        f.close()

    def writefile(eno,pdfwrite):
        """  Function to create blank PDF file and Copy pages into it       """
        fname = '{}.pdf'.format(eno)
        f = os.path.join(args.targetfiledirectory,fname.replace(' ',''))
        targetfile = open(f,'wb')
        pdfwrite.write(targetfile)
        targetfile.close()

    def splitonstr(pdfread,data):
        """  Function to Split the PDF file by key value of dictionary object       """
        for stfno in data:
            log('Employee {} - No of Pages {}'.format(stfno,len(data[stfno])))
            pdfwriter = PDF.PdfFileWriter()
            pdfwriter.removeImages()
            for page in data[stfno]:
                pageobj = pdfread.getPage(page)
                pdfwriter.addPage(pageobj)
            writefile(stfno,pdfwriter) 

    """ Main() First parse the argument """   

    argps = argparse.ArgumentParser()
    argps.add_argument("--sourcefiledirectory", help="File Directory", required=True)
    argps.add_argument("--targetfiledirectory", help="File Directory", required=True)
    args = argps.parse_args()
    totalemp = 0
    # Create log file for storing the Key and Pages 

    flog = os.path.join(args.sourcefiledirectory,"Employee_Paylog" + time.strftime('%Y%m%d%H%M%S')+".txt")

    # Get PDF file list from the source directory and exit if no file exist

    files = [f for f in os.listdir(args.sourcefiledirectory) if f.upper().endswith(".PDF")] 
    if len(files) == 0:
        log('No File To Process')
        exit(1)

    # Compress the source PDF file
    log('Compressing Source PDF file Using Ghost Script')
    for f in files :
        pdfout = os.path.join(args.sourcefiledirectory,'cmp'+f)
        pdfin = os.path.join(args.sourcefiledirectory,f)
        compresspdf(pdfout,pdfin)

    log('Compressing Completed')

    
    #Iterate through each PDF file in Source and Split for each employee

    for f in files :
        log("Processing Files {}".format(f))
        pdffilepath = os.path.join(args.sourcefiledirectory,f)
        pdffile = open(pdffilepath,'rb')
        pdfreader = PDF.PdfFileReader(pdffile)
        data = getstaffno_page(pdfreader)
        writeemployeepaylog(flog,data)
        totalemp = totalemp + len(data)
        splitonstr(pdfreader,data)

        log('Creating Archive For Split Files')
        zfiles = [os.path.join(args.targetfiledirectory, f) for f in os.listdir(args.targetfiledirectory) if not os.path.isdir(os.path.join(args.targetfiledirectory, f)) and f.upper().endswith(".PDF") and f.upper() not in files]
        zipfname = os.path.join(args.targetfiledirectory,'Payslip_'+time.strftime('%Y%m%d%H%M%S')+'.zip')

        if len(zfiles):
            makeArchive(zfiles, zipfname, os.path.join(args.targetfiledirectory))
        log('Archive Created Successfully')
    pdffile.close()
    
    log('All PDF files processed successfully')

    log('Renaming and Moving PDF Source files into Archive directory')
    for f in files:
    	newname = time.strftime('%Y%m%d%H%M%S') +'_'+f
    	shutil.move(os.path.join(args.sourcefiledirectory,f),os.path.join(args.sourcefiledirectory,'Archive',newname))

    log('Start Processing Manifest File')
    """ Process the Manifest file provided by chris 21 to :
            1.  Remove the extra header i.e top row of the CSV file and make the 2nd row the header
            2.  Convert the date fields to date time with UTC offset
            3.  Create Column File_Name which is the Name of the Split PDF and concatenation of Employee Number , Pay run and Pay Date
    """
    tzdelta = '+10:00' # datetimedelta between current local time and UTC time
     
    files =  [f for f in os.listdir(args.sourcefiledirectory) if not os.path.isdir(os.path.join(args.sourcefiledirectory, f)) and f.upper().endswith(".CSV")]
    if len(files) > 0 :
        for f in files:
            df = pd.read_csv(os.path.join(args.sourcefiledirectory, f))
            df.columns = list(df.loc[0].values)
            df.drop(0, inplace = True, axis = 0)
            """ Create FileName """
            df["File_Name"] = df['Worker'] + '_' + df['Check_Number'] + '_' + df['Display_Date'].apply(lambda x : x.replace(' ','')+'.pdf')
            df['Pay_Period_End_Date'] = pd.to_datetime(df['Pay_Period_End_Date'])
            df['Payment_Date'] =  pd.to_datetime(df['Payment_Date'])
            df['Display_Date'] =  pd.to_datetime(df['Display_Date'])
            df['Pay_Period_End_Date'] = df['Pay_Period_End_Date'].apply(lambda x : (str(x)[0:11]+tzdelta).replace(' ',''))
            df['Payment_Date'] =  df['Payment_Date'].apply(lambda x : (str(x)[0:11]+tzdelta).replace(' ',''))
            df['Display_Date'] =  df['Display_Date'].apply(lambda x : (str(x)[0:11]+tzdelta).replace(' ',''))
            newname = time.strftime('%Y%m%d%H%M%S') +'_'+f
            shutil.move(os.path.join(args.sourcefiledirectory,f),os.path.join(args.sourcefiledirectory,'Archive',newname))
            df.to_csv(os.path.join(args.targetfiledirectory, f),index=False )

    log('All process completed Successfully')
    log('Summary : Total Employee IDs Processed = {}'.format(totalemp))

except Exception as e :
	logging.exception("message")
	log('Process did not complete successfully')
	exit(1)