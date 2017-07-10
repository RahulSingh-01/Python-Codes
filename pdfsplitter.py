import PyPDF2 as PDF
import os
import argparse
import logging
import time
import shutil
import zipfile
import re

#python pdfsplitter.py --sourcefiledirectory "C:\\Temp\PDF\Source" --targetfiledirectory "C:\\Temp\PDF\Target"

try :

    def create_logger():
        logging.basicConfig(filename='pdfsplitter.log',level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    def log(msg):
        logging.info(msg)
        print('[' + time.strftime('%Y-%m-%d %H:%M:%S') + '] ' + msg)

    create_logger()

    argps = argparse.ArgumentParser()
    argps.add_argument("--sourcefiledirectory", help="File Directory", required=True)
    argps.add_argument("--targetfiledirectory", help="File Directory", required=True)
    args = argps.parse_args()

    files = [f for f in os.listdir(args.sourcefiledirectory) if f.upper().endswith(".PDF")] 

    if len(files) == 0:
        log('No File To Process')
        exit(1)
    for f in files :
        log("Processing Files {}".format(f))
        pdffilepath = os.path.join(args.sourcefiledirectory,f)
        pdffile = open(pdffilepath,'rb')
        pdfreader = PDF.PdfFileReader(pdffile)
        x = pdfreader.numPages
        stfnum = []
        for i in range(x):
            pageobj = pdfreader.getPage(i)
            a = pageobj.extractText()
            m = re.search('Staff:(.+?) ',a)
            stfnum.append(m.group(1).rstrip().lstrip())
        pdfwrite = PDF.PdfFileWriter()
        emp = stfnum[0]
        stfnum.append('lastemp')
        for i in range(len(stfnum)):
            if emp == stfnum[i]:
                pageobj = pdfreader.getPage(i)
                pdfwrite.addPage(pageobj)
            else :
                fname = '{}.pdf'.format(emp)
                tfpdfpath = os.path.join(args.targetfiledirectory,fname)
                targetfile = open(tfpdfpath,'wb')
                pdfwrite.write(targetfile)
                targetfile.close()
                if stfnum[i] != 'lastemp':
                    emp = stfnum[i]
                    pdfwrite = PDF.PdfFileWriter()
                    pageobj = pdfreader.getPage(i)
                    pdfwrite.addPage(pageobj)
        zfiles = [os.path.join(args.targetfiledirectory, f) for f in os.listdir(args.targetfiledirectory) if not os.path.isdir(os.path.join(args.targetfiledirectory, f)) and f.upper().endswith(".PDF") and f.upper() not in files]
        if len(zfiles):
            zipfname = 'Payslip_'+time.strftime('%Y%m%d%H%M%S')+'.zip'
            zf = os.path.join(args.targetfiledirectory,zipfname)
            fantasy_zip = zipfile.ZipFile(zf, 'w')
            for f in zfiles:
                if f.upper().endswith('.PDF'):
                    fantasy_zip.write(os.path.join(args.targetfiledirectory,f), compress_type = zipfile.ZIP_DEFLATED)
                    os.remove(os.path.join(args.targetfiledirectory,f))
            fantasy_zip.close()
        pdffile.close()

    log('All files processed successfully')

    log('Moving Source files into Archive')
    for f in files:
    	#log('Processing file {}'.format(f))
    	newname = time.strftime('%Y%m%d%H%M%S') +'_'+f
    	shutil.move(os.path.join(args.sourcefiledirectory,f),os.path.join(args.sourcefiledirectory,'Archive',newname))

    log('All process completed Successfully')

except Exception as e :
	logging.exception("message")
	log('Process did not complete successfully')
	exit(1)