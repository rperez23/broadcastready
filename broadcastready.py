#! /usr/bin/python3

import openpyxl
import os
import pandas as pd
import re
import sys
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

#Conditions for broadcast ready
#1: mxf
#2: hn assigned
#3: hn.scc file on s3

wslist = ['Full Ingest Summary', 'Captions Summary']
housenumbers = {}


def usage():
	print('')
	print('Usage:')
	print('')
	print('   broadcastready.py <xlf name>')
	print('')

def getxlf():

	xlf = sys.argv[1]

	if os.path.isfile(xlf):
		return xlf

	print('')
	print(xlf,'Does not exist, exiting')
	usage()
	sys.exit(1)

def gethousenumbers():

	hnlist = {}

	print('')
	print('ENTER YOUR HOUSE NUMBERS:')
	print('')

	while(1):

		hn = input()

		if hn == '':
			break

		m  = re.match('^BUZ_[A-Z0-9]+$',hn)

		if m:
			hnlist[hn] = '1'


	#print(hnlist)


if len(sys.argv) <= 1:
	usage()
	sys.exit(1)


xlf = getxlf()
housenumbers = gethousenumbers()

#open the workbook for reading
try:
	wb = openpyxl.load_workbook(filename=xlf,read_only=True)
except:
	print("  Cannot open",xlf)
	sys.exit(1)

#validate the worksheets are in the xlf
#validate that sheet exists
for tab in wslist:
	if not (tab in wb.sheetnames):
		print(" ",tab,"not in",xlf)
		print('')
		wb.close()
		sys.exit(1)
wb.close()


videodb   = {}
captiondb = {}

df = pd.read_excel(xlf,sheet_name=wslist[0])
videodb = df.to_dict()

df = pd.read_excel(xlf,sheet_name=wslist[1])
captiondb = df.to_dict()






