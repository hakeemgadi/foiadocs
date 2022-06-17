import pytesseract
from PIL import Image
import cv2
import numpy as np
import pdf2image
import os
import PyPDF2
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.converter import TextConverter
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfdocument import PDFDocument


from pdfminer.layout import LAParams, LTTextBox, LTTextLine

from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

from io import StringIO
import unicodedata
from difflib import SequenceMatcher

def remove_duplicate_lines(txt):
    linelst=txt.split("\n")
    newlinelst=[]
    for inx,line in enumerate(linelst[0:-1]):
        if SequenceMatcher(None,line,linelst[inx+1]).ratio()<0.85:
            newlinelst.append(line)
    return("\n".join(newlinelst))
            
        

def extract_txt3(pfile,pg):
    fp = open(pfile, 'rb')
    txt = ''
    parser = PDFParser(fp)
    doc = PDFDocument(parser,'')
    parser.set_document(doc)
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    # Process each page contained in the document.
    pages=list(PDFPage.create_pages(doc))
    interpreter.process_page(pages[pg-1])
    layout = device.get_result()
    for lt_obj in layout:
        if isinstance(lt_obj, LTTextBox):
            txt += lt_obj.get_text()
    return(remove_duplicate_lines(txt)) 

def extract_txt2(pfile):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(pfile, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return(remove_duplicate_lines(text))

def extract_txt(pfile,pg):
    pdfFileObj = open(pfile, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    pdfReader.numPages
    pageObj = pdfReader.getPage(pg-1)
    txt=pageObj.extractText()
    print("========Page {} extract=========".format(pg))
    pdfFileObj.close()
    return(txt)

def ocr_txt(pfile,pg):
    page = pdf2image.convert_from_path(pfile, dpi=400,fmt='png',first_page=pg,last_page=pg)[0]
    I = np.array(page)
    image = cv2.resize(I, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    retval, threshold = cv2.threshold(image,127,255,cv2.THRESH_BINARY)

    text = pytesseract.image_to_string(image)
    return(text)
    

pdflst=os.listdir(os.path.join(".","PDFS"))
pdflst.sort()

print(len(pdflst))
startfl=1839
for ix,pdf in enumerate(pdflst[startfl-1:-1]):
    print("###{}### Working on PDF {}...\n".format(startfl+ix,pdf))
    pdf_file=os.path.join(".","PDFS",pdf)
    
    info = pdf2image.pdfinfo_from_path(pdf_file)
    


    numpages=info["Pages"]
    startpg=1#Normally 1
    for idx in range(numpages):
        currpg=idx+startpg
        
        print("<<<<<<<<<<<<<< In file number: {} >>>>>>>>>>>>>>".format(startfl+ix))
        #Work one page at a time or code will eat up memory fast
        print("Extracting page {}/{}...\n".format(currpg,numpages))
        text=extract_txt3(pdf_file,currpg)
        if len(text)<10:
            print("PAGE SEEMS TO BE IMAGE")
        headlen=200
        print("First {} characters of EXT text:\n".format(headlen))
        print(text[0:headlen],"\n")
        #Save
        file = open(os.path.join("..","DocStore","TXTEXT","{}-{:03d}-EXT.txt".format(pdf.split(".")[0],currpg)),"w",encoding="utf-8")
        file.write(text)
        file.close()


        print("Converting page {}/{}...\n".format(currpg,numpages))

        text=ocr_txt(pdf_file,currpg)     

        print("First {} characters of OCR text:\n".format(headlen))
        print(text[0:headlen],"\n")
        #Save
        file = open(os.path.join("..","DocStore","TXTOCR","{}-{:03d}-OCR.txt".format(pdf.split(".")[0],currpg)),"w",encoding="utf-8")
        file.write(text)
        file.close()
