#! /usr/bin/python3

import pandas as pd
import re
from urllib.request import urlretrieve
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

xlf = 'Broadcast-Ready-Status.xlsx'
tab = 'Sheet'

datareport = pd.read_excel(xlf,tab,usecols='A:F')

for i in range(0,len(datareport)):

	capexists = datareport['S3 HN CAPTION'][i]
	capurl    = datareport['ASSET CAPTION'][i]
	hn        = datareport['HOUSE NUMBER'][i]

	if (capexists == 'NO') and (str(capurl) != 'nan'):

		print('downloading:',hn)
		print(capurl)

		ext = '.scc'
		msrt = re.search('\.srt?',capurl)

		if msrt:
			ext = '.srt'


		fname = hn + ext

		urlretrieve(capurl,fname)

