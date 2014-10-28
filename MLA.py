#TODO
#Indent --Unlikely
#Quote Format
#Fix header selection
#Methodize everything
#Share Vars between classes
#COMPLETE
#Input Info
#Font
#Page Number
#Margins
#Last Name By Page Number
#
#
#Select File
import Tkinter, tkFileDialog
from Tkinter import *
from tkMessageBox import *

#Arrays of Found Text Indices
fontocc = []
rightt = []
doublespaced = []
pgn = []
lastnamehdr = []
#MLA format true/false
isMLA = True

   
root = Tkinter.Tk()
root.withdraw()

filepath = tkFileDialog.askopenfilename()
print filepath
#Finder function
def finder(pattern, text, fun):
   for match in re.finditer(pattern, text):
      s = match.start()
      e = match.end()
      fun.append(e)
      print 'Found "%s" at %d:%d' % (text[s:e], s, e)\
#Zip import
import zipfile
#Get the header and document XML
def get_word_xml(docx_filename, fil):
   with file(docx_filename, 'rb') as f:
      zip = zipfile.ZipFile(f)
      xml_content = zip.read('word/' + fil + '.xml')
   return xml_content
#etree import
from lxml import etree
#Get xml Tree
def get_xml_tree(xml_string):
   return etree.fromstring(xml_string)
#Document xml and header1 xml variables
tree = get_xml_tree(get_word_xml(filepath, 'document'))
hed = 1
def correctheader(num):
   try:
      global hed
      pgnum = get_xml_tree(get_word_xml(filepath, 'header'+str(num)))
      print num
      if '<w:t>' in etree.tostring(get_xml_tree(get_word_xml(filepath, 'header'+str(num)))):
         print 'Header ' + str(num)
         hed = num
         return
      else:
         correctheader(num+1)
   except KeyError:
         print 'No Header with Content found'
         return
correctheader(1)
print hed
pgnum = get_xml_tree(get_word_xml(filepath, 'header'+str(hed)))
#Dump main xml doc to console and convert the xml into usable strings
#print etree.tostring(tree, pretty_print=True)
text = etree.tostring(tree)
numm = etree.tostring(pgnum)
#Regex import
import re


#Detail Input
name = raw_input('Input Name: ')
teacher = raw_input('Input Teacher')
clas = raw_input('Input Class in format: Class/period')
date = raw_input('Input Date in format: Numerical Day of Month Name of Month Numerical Year')
#Detail Input Placeholders
#name = 'Anders Sundheim'
#teacher = 'Ms. Downey'
#clas = 'English I/5'
#date = '8 October 2014'




#Import Document Class
from docx import Document
from docx.shared import Inches
#Define Document
doc = Document(filepath)
sec = doc.sections
secc = sec[0]
par = doc.paragraphs
#Margins

#Header/Footer
def number():
   global numm
   global rightt
   global isMLA
   #Find Page Number
   finder('right', numm, rightt)
   finder('Page Numbers', numm, pgn)
   if len(rightt)>1 and len(pgn)>0:
      print 'Page Number Correct'
      showinfo('Page Number', 'Page Number Found')
   else:
      isMLA = False
      print 'No Page Number Found On Right Header'
#
#
#
#Double Spaced
def doublespace():
   global isMLA
   global tree
   global doublespaced
   global text
   doub = True
   finder('w:line=', etree.tostring(tree), doublespaced)
   for i in range(len(doublespaced)):
      if int(text[doublespaced[i]+1:doublespaced[i]+4]) == 480:
         #print 'Doublespace On Line ' + str(i+1)
         isMLA = isMLA         
      else:
         print int(text[doublespaced[i]+1:doublespaced[i]+4])
         print 'No Double Space on Line ' + str(i+1)
         isMLA = False
         doub = False
   showinfo('Double Spaced', 'Entire Document Double Spaced = ' + str(doub))
#
#
#
def header():
   global secc
   global isMLA
   if secc.header_distance.inches == 0.5 and secc.footer_distance.inches == 0.5:
      print 'Headers Good'
      showinfo('Headers', 'Headers Good')
   else:
      print 'Headers Bad'
      print 'Header: ' +  str(secc.header_distance.inches)
      print 'Footer: ' + str(secc.footer_distance.inches)
      isMLA = False
      showinfo('Margins', 'Margins Bad')
#
#
#
#Top 4 Lines Checking
def toplines():
   global isMLA
   global par
   global teacher
   global clas
   global name
   global date
   if par[0].text == name and par[1].text == teacher and (par[2].text == clas or par[2].text == clas + 'th') and par[3].text == date:
      print 'Top 4 lines OK'
      showinfo('Top 4 Lines', 'Top 4 Lines Ok')
   else:
      #print 'name ' + par[0].text == name
      #print 'teacher ' + par[1].text == teacher
      print 'Top 4 lines not OK'
      isMLA = False
      showinfo('Top 4 Lines', 'Top 4 Lines Not Ok')
#
#
#
#Margins
def margins():
   global isMLA
   global secc   
   if secc.left_margin.inches == 1 and secc.right_margin.inches == 1 and secc.bottom_margin.inches == 1 and secc.top_margin.inches == 1:
      print 'Margins Good'
      showinfo('Margins', 'Margins Good')
   else:
      print 'Margins Bad'
      print 'Left margin: ' + str(secc.left_margin.inches)
      print 'Right margin: ' + str(secc.right_margin.inches)
      print 'Top margin: ' + str(secc.top_margin.inches)
      print 'Bottom margin: ' + str(secc.bottom_margin.inches)
      isMLA = False
      showinfo('Margins', 'Margins Bad')
#
#
#
#Find Font
def font():
   global tree
   global fontocc
   global text
   global isMLA
   finder('Font', etree.tostring(tree), fontocc)
   #Font
   for i in range(len(fontocc)):
      if text[fontocc[i]+11] == 'T' or (text[fontocc[i]+4:fontocc[i]+8] == 'hint'):
         #print 'Font Times New Roman at paragraph ' + str(i+1)
         isMLA = isMLA
         showinfo('Font Correct')
      else:
         print text[fontocc[i]+11]
         print 'bad'
         showinfo('Font', 'Font Returned False, If Incorrect Disregard')
         isMLA = False
#
#
#
def lastname_header():
   global isMLA
   global lastnamehdr
   global pgnum
   global numm
   #Check For Last Name on Header
   finder('<w:t>', etree.tostring(pgnum), lastnamehdr)
   lastname = name.split(' ')
   print lastname
   if lastname[1] in numm:#str(numm[lastnamehdr[0]:lastnamehdr[0]+len(lastname[1])]) == lastname[1]:
      print 'Last Name Header OK'
      showinfo('Last Name on Header', 'Last Name In Header = True')
   else:
      print str(numm[lastnamehdr[0]:lastnamehdr[0]+len(lastname[1])])
      print lastname[1]
      print str(numm[lastnamehdr[0]:lastnamehdr[0]+len(lastname[1])])==lastname[1]
      print 'No Last Name Preceding Header'
      isMLA = False
      showinfo('Last Name on Header', 'No Last Name In Header')
#
#
#
#Call Checker Functions
margins()
header()
toplines()
doublespace()
font()
lastname_header()
#Result
print 'MLA = ' + str(isMLA)
showinfo('MLA', 'MLA = ' + str(isMLA))
