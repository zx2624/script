from __future__ import division 
from openpyxl import Workbook
from openpyxl import load_workbook
import os
import sys

print sys.path[0]
files=os.listdir(sys.path[0])
for file in files:
	s=file.split(".")
	print s[0]

	
