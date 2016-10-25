import os
import sys
import re
import argparse
from datetime import datetime

import xlrd
import validators
import gspread

from oauth2client.service_account import ServiceAccountCredentials
from ebaysdk.trading import Connection as Trading
from ebaysdk.exception import ConnectionError

class EbayPublisher(object):

	def read_sheet(self, _file):
		book = xlrd.open_workbook(_file)
		sh = book.sheet_by_index(0)
		all_values = [sh.row_values(x) for x in range(sh.nrows)]
		self.process_data(all_values)

	def read_gsheet(self, _url):
		credentials = ServiceAccountCredentials.from_json_keyfile_name('gsheets_credentials.json', 'https://spreadsheets.google.com/feeds')
		gc = gspread.authorize(credentials)
		book = gc.open_by_url(_url)
		sh = book.get_worksheet(0)
		all_values = sh.get_all_values()
		self.process_data(all_values)

	def process_data(self, values):
		for index, row in enumerate(values):
			if index == 0:
				continue

			item = self.format_item(index, row)
			resp = self.add_item(item)

			if resp:
				print 'Item insertion success:'
				print 'ItemID: {}'.format(resp.get('ItemID'))
				print '------------'

	def add_item(self, item):
		try:
			api = Trading(domain='api.sandbox.ebay.com', config_file='api_config.yaml')
			return api.execute('VerifyAddItem', item).dict() # change for AddItem resource
		except ConnectionError as e:
			print 'Item insertion failed: \n\n'
			print e.response.dict()

	def format_item(self, index, values):
		error_fields = []

		if not values[0]:
			error_fields.append('SKU')
		if not values[1]:
			error_fields.append('UPC')
		if not values[2]:
			error_fields.append('Cost')
		if not values[3]:
			error_fields.append('Title')
		if not values[4]:
			error_fields.append('Description')
		if not values[5]:
			error_fields.append('Category')

		if len(error_fields) > 0:
			print 'Error: Line \'{}\'. Columns \'{}\' should not be null: '.format(str(index+1), ', '.join(error_fields))            
			sys.exit()
		
		return {
		    "Item": {
				"SKU": values[0],
				"ProductListingDetails": {
				   "UPC": str(values[1]),
				},
				"StartPrice": values[2],
				"Title": self.clean(values[3]),
				"Description": self.clean(values[4]),
				"PrimaryCategory": {
					"CategoryID": int(values[5])
				},
				"PictureDetails": {
				    "PictureURL": filter(None, [values[6], values[7], values[8], values[9], values[10], values[11]])
				},
				"ItemSpecifics": {
				    "NameValueList" : [{
		                "Name": '' if len(values[12].split('|')) == 1 else self.clean(values[12].split('|')[0]), 
						"Value": '' if len(values[12].split('|')) == 1 else self.clean(values[12].split('|')[1]) 
				    }, {
						"Name": '' if len(values[13].split('|')) == 1 else self.clean(values[13].split('|')[0]),
						"Value": '' if len(values[13].split('|')) == 1 else self.clean(values[13].split('|')[1])
				    }, {
						"Name": '' if len(values[14].split('|')) == 1 else self.clean(values[14].split('|')[0]),
						"Value": '' if len(values[14].split('|')) == 1 else self.clean(values[14].split('|')[1])
				    }, {
						"Name": '' if len(values[15].split('|')) == 1 else self.clean(values[15].split('|')[0]),
						"Value": '' if len(values[15].split('|')) == 1 else self.clean(values[15].split('|')[1])
				    }, {
						"Name": '' if len(values[16].split('|')) == 1 else self.clean(values[16].split('|')[0]),
						"Value": '' if len(values[16].split('|')) == 1 else self.clean(values[16].split('|')[1])
				    }, {
						"Name": '' if len(values[17].split('|')) == 1 else self.clean(values[17].split('|')[0]),
						"Value": '' if len(values[17].split('|')) == 1 else self.clean(values[17].split('|')[1])
				    }, {
						"Name": '' if len(values[18].split('|')) == 1 else self.clean(values[18].split('|')[0]),
						"Value": '' if len(values[18].split('|')) == 1 else self.clean(values[18].split('|')[1])
				    }, {
				        "Name": '' if len(values[19].split('|')) == 1 else self.clean(values[19].split('|')[0]),
						"Value": '' if len(values[19].split('|')) == 1 else self.clean(values[19].split('|')[1])
				    }, {
						"Name": '' if len(values[20].split('|')) == 1 else self.clean(values[20].split('|')[0]),
						"Value": '' if len(values[20].split('|')) == 1 else self.clean(values[20].split('|')[1])
				    }, {
						"Name": '' if len(values[21].split('|')) == 1 else self.clean(values[21].split('|')[0]),
						"Value": '' if len(values[21].split('|')) == 1 else self.clean(values[21].split('|')[1])
				    }, {
						"Name": '' if len(values[22].split('|')) == 1 else self.clean(values[22].split('|')[0]),
						"Value": '' if len(values[22].split('|')) == 1 else self.clean(values[22].split('|')[1])
				    }, {
						"Name": '' if len(values[23].split('|')) == 1 else self.clean(values[23].split('|')[0]),
						"Value": '' if len(values[23].split('|')) == 1 else self.clean(values[23].split('|')[1])
				    }, {
						"Name": '' if len(values[24].split('|')) == 1 else self.clean(values[24].split('|')[0]),
						"Value": '' if len(values[24].split('|')) == 1 else self.clean(values[24].split('|')[1])
				    }, {
						"Name": '' if len(values[25].split('|')) == 1 else self.clean(values[25].split('|')[0]),
						"Value": '' if len(values[25].split('|')) == 1 else self.clean(values[25].split('|')[1])
				    }, {
						"Name": '' if len(values[26].split('|')) == 1 else self.clean(values[26].split('|')[0]),
						"Value": '' if len(values[26].split('|')) == 1 else self.clean(values[26].split('|')[1])
				    }]
				},
				"SellerProfiles":{
					"SellerPaymentProfile": {
						"PaymentProfileID": "60182845013"
					},
					"SellerReturnProfile":{
						"ReturnProfileID": "63713325013"
					},
					"SellerShippingProfile":{
						"ShippingProfileID": "77108635013"
					}	
				},

				# Mandatory fields
				"CategoryMappingAllowed": "true",
				"Country": "US",
				"Currency": "USD",
				"ConditionID": "3000",
				"DispatchTimeMax": "3",
				"ListingDuration": "Days_7",
				"PostalCode": "95125",
				
		    }
		}

	def clean(self, _str):
		return re.sub(r'&([^a-zA-Z#])',r'&amp;\1', _str)


def get_date():
    dt = datetime.now()
    return dt.strftime('%Y-%m-%d %H:%M:%S')

def run_main(args):
    
    ep = EbayPublisher()
    if args.url:
		if not validators.url(args.url):
			print '{} ERROR:  {} is not a valid google sheet url.'.format(get_date(), args.file)
			sys.exit() 
		ep.read_gsheet(args.url)
    
    else:
		if not os.path.isfile(args.file):
			print '{} ERROR:  {} is not a valid file.'.format(get_date(), args.file)
			sys.exit() 
		ep.read_sheet(args.file)

if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument("-file", help='The file path you want to read')
    parser.add_argument("-url", help='Specifies if the file is a Google Sheet')
    args = parser.parse_args()

    try:
		run_main(args)
    except Exception as e:
    	s = str(e)
	print '{} ERROR:  {}'.format(get_date(), s)
