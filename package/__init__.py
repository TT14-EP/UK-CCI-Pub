__version__ = '0.1'
__author__ = 'Thomas Tang'
__email__ = 'thomas.tang@evelyn.com'

print("Initializing package...")

# Import libraries
from docxtpl import DocxTemplate
from docx.shared import Inches
from docx import Document
from docx2pdf import convert

import pyodbc
import pandas as pd
import os
import glob
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
import datetime
import re
import sys
import babel.numbers


print('===================================================================================================================')
print('Greetings Earthling. You are accessing the CCI publication module.')
print('Please check paths. You may amend the file names of Python datasheet and Word template in next step.')
print('===================================================================================================================')
print('Version: ' + __version__)
print('Author: ' + __author__)
print('Email: ' + __email__)
print('\n')

# Set out paths for folders, word template and data sheet
wdtemplatename = 'CCI Product Summary Template v0.4.docx'
datasheetname = '202605_PyDataSheet_260608150009.csv'

plotfolderpath = os.path.abspath(os.path.join(os.getcwd(), 'plotdump'))
pdffolderpath = os.path.abspath(os.path.join(os.getcwd(), 'output'))
wdtemplatepath = os.path.abspath(os.path.join(os.getcwd(), wdtemplatename))
datasheetpath = os.path.abspath(os.path.join(os.getcwd(), datasheetname))

print('Plots folder: ' + plotfolderpath)
print('PDF folder: ' + pdffolderpath)
print('Word template: ' + wdtemplatepath)
print('Python datasheet: ' + datasheetpath)
print('===================================================================================================================\n')