#! /usr/bin/python3

import openpyxl
import os
import sys
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

#Conditions for broadcast ready
#1: mxf
#2: hn assigned
#3: hn.scc file on s3


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



if len(sys.argv) <= 1:
	usage()
	sys.exit(1)


getxlf()
