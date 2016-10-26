import os
import sys
import re
import json
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

        try:
            book = gc.open_by_url(_url)
        except:
            print 'ERROR: SpreadsheetNotFound'
            sys.exit()

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
        with open('api_config.json') as econfig:
            config = json.load(econfig)
        try:
            api = Trading(domain=config.get('domain'), appid=config.get('appid'), devid=config.get('devid'), certid=config.get('certid'), token=config.get('token'), config_file=None)
            return api.execute('VerifyAddItem', item).dict() # change for AddItem resource
        except ConnectionError as e:
            print 'Item insertion failed: \n\n'
            print e.response.dict()

    def format_item(self, index, values):
        error_fields = []

        """
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
        """

        if len(error_fields) > 0:
            print 'Error: Line \'{}\'. Columns \'{}\' should not be null: '.format(str(index+1), ', '.join(error_fields))            
            sys.exit()
        
        return {
            "Item": {
                "DiscountPriceInfo":{
                    "MinimumAdvertisedPrice": values[12]
                },

                "SKU": values[2],
                "ProductListingDetails": {
                   "UPC": str(values[6]),
                },
                "StartPrice": values[13],
                "Title": self.clean(values[14]),
                "Description": self.clean(values[4]),
                "PrimaryCategory": {
                    "CategoryID": int(values[15])
                },
                "PictureDetails": {
                    "PictureURL": filter(None, [values[16], values[17], values[18], values[19], values[20], values[21]])
                },
                "ItemSpecifics": {
                    "NameValueList" : [{
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
                    }, {
                        "Name": '' if len(values[27].split('|')) == 1 else self.clean(values[27].split('|')[0]),
                        "Value": '' if len(values[27].split('|')) == 1 else self.clean(values[27].split('|')[1])
                    }, {
                        "Name": '' if len(values[28].split('|')) == 1 else self.clean(values[28].split('|')[0]),
                        "Value": '' if len(values[28].split('|')) == 1 else self.clean(values[28].split('|')[1])
                    }, {
                        "Name": '' if len(values[29].split('|')) == 1 else self.clean(values[29].split('|')[0]),
                        "Value": '' if len(values[29].split('|')) == 1 else self.clean(values[29].split('|')[1])
                    }, {
                        "Name": '' if len(values[30].split('|')) == 1 else self.clean(values[30].split('|')[0]),
                        "Value": '' if len(values[30].split('|')) == 1 else self.clean(values[30].split('|')[1])
                    }, {
                        "Name": '' if len(values[31].split('|')) == 1 else self.clean(values[31].split('|')[0]),
                        "Value": '' if len(values[31].split('|')) == 1 else self.clean(values[31].split('|')[1])
                    }, {
                        "Name": '' if len(values[32].split('|')) == 1 else self.clean(values[32].split('|')[0]),
                        "Value": '' if len(values[32].split('|')) == 1 else self.clean(values[32].split('|')[1])
                    },{
                        "Name": '' if len(values[33].split('|')) == 1 else self.clean(values[33].split('|')[0]), 
                        "Value": '' if len(values[33].split('|')) == 1 else self.clean(values[33].split('|')[1]) 
                    }, 
                    {
                        "Name": '' if len(values[34].split('|')) == 1 else self.clean(values[34].split('|')[0]),
                        "Value": '' if len(values[34].split('|')) == 1 else self.clean(values[34].split('|')[1])
                    }, {
                        "Name": '' if len(values[35].split('|')) == 1 else self.clean(values[35].split('|')[0]),
                        "Value": '' if len(values[35].split('|')) == 1 else self.clean(values[35].split('|')[1])
                    }, {
                        "Name": '' if len(values[36].split('|')) == 1 else self.clean(values[36].split('|')[0]),
                        "Value": '' if len(values[36].split('|')) == 1 else self.clean(values[36].split('|')[1])
                    }, 
                    
                    ]
                },
                "SellerProfiles":{
                    "SellerPaymentProfile": {
                        "PaymentProfileID": "60182845013"
                        #"PaymentProfileID": "5452208000"
                    },
                    "SellerReturnProfile":{
                        "ReturnProfileID": "63713325013"
                        #"ReturnProfileID": "5451872000"
                    },
                    "SellerShippingProfile":{
                        "ShippingProfileID": "77108635013"
                        #"ShippingProfileID": "5452249000"
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
