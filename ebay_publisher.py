import os
import xlrd
import argparse
from datetime import datetime

from ebaysdk.trading import Connection as Trading
from ebaysdk.exception import ConnectionError

class EbayPublisher(object):

    def process_data(self, _file):

	book = xlrd.open_workbook(_file)
	sh = book.sheet_by_index(0)
	
        for row in range(sh.nrows):
	    if row == 0:
		continue

	    values = sh.row_values(row)
   	    #import ipdb; ipdb.set_trace()
	    item = {
		"Item": {
		    "SKU": values[0],
		    #"ProductListingDetails": {
		    # 	"UPC": str(values[1]),
		    #},
		    "StartPrice": values[2],
		    "Title": values[3],
		    "Description": values[4],
		    "PrimaryCategory": {
                        "CategoryID": "377"#int(values[5])
                    },
                    "PictureDetails": {
			"PictureURL": values[6]
		    },
		    "ItemSpecifics": {
		        "NameValueList" : [{
			    	    "Name": '' if len(values[12].split('|')) == 1 else values[12].split('|')[0], 
			    	    "Value": '' if len(values[12].split('|')) == 1 else values[12].split('|')[1] 
				}, {
				    "Name": '' if len(values[13].split('|')) == 1 else values[13].split('|')[0],
				    "Value": '' if len(values[13].split('|')) == 1 else values[13].split('|')[1]
				}, {
				    "Name": '' if len(values[14].split('|')) == 1 else values[14].split('|')[0],
				    "Value": '' if len(values[14].split('|')) == 1 else values[14].split('|')[1]
				}, {
				    "Name": '' if len(values[15].split('|')) == 1 else values[15].split('|')[0],
				    "Value": '' if len(values[15].split('|')) == 1 else values[15].split('|')[1]
				}, {
				    "Name": '' if len(values[16].split('|')) == 1 else values[16].split('|')[0],
				    "Value": '' if len(values[16].split('|')) == 1 else values[16].split('|')[1]
				}, {
				    "Name": '' if len(values[17].split('|')) == 1 else values[17].split('|')[0],
				    "Value": '' if len(values[17].split('|')) == 1 else values[17].split('|')[1]
				}, {
				    "Name": '' if len(values[18].split('|')) == 1 else values[18].split('|')[0],
				    "Value": '' if len(values[18].split('|')) == 1 else values[18].split('|')[1]
				}, {
				    "Name": '' if len(values[19].split('|')) == 1 else values[19].split('|')[0],
				    "Value": '' if len(values[19].split('|')) == 1 else values[19].split('|')[1]
				}, {
				    "Name": '' if len(values[20].split('|')) == 1 else values[20].split('|')[0],
				    "Value": '' if len(values[20].split('|')) == 1 else values[20].split('|')[1]
				}, {
				    "Name": '' if len(values[21].split('|')) == 1 else values[21].split('|')[0],
				    "Value": '' if len(values[21].split('|')) == 1 else values[21].split('|')[1]
				}, {
				    "Name": '' if len(values[22].split('|')) == 1 else values[22].split('|')[0],
				    "Value": '' if len(values[22].split('|')) == 1 else values[22].split('|')[1]
				}, {
				    "Name": '' if len(values[23].split('|')) == 1 else values[23].split('|')[0],
				    "Value": '' if len(values[23].split('|')) == 1 else values[23].split('|')[1]
				}, {
				    "Name": '' if len(values[24].split('|')) == 1 else values[24].split('|')[0],
				    "Value": '' if len(values[24].split('|')) == 1 else values[24].split('|')[1]
				}, {
				    "Name": '' if len(values[25].split('|')) == 1 else values[25].split('|')[0],
				    "Value": '' if len(values[25].split('|')) == 1 else values[25].split('|')[1]
				}, {
				    "Name": '' if len(values[26].split('|')) == 1 else values[26].split('|')[0],
				    "Value": '' if len(values[26].split('|')) == 1 else values[26].split('|')[1]
				}]
		    },

                    # Mandatory fields
		    "CategoryMappingAllowed": "true",
		    "Country": "US",
		    "Currency": "USD",
		    "ConditionID": "3000",
		    "DispatchTimeMax": "3",
		    "ListingDuration": "Days_7",
		    "PaymentMethods": "PayPal",
		    "PayPalEmailAddress": "gleontra@gmail.com",
		    "PostalCode": "95125",
		    "ReturnPolicy": {
		    	"ReturnsAcceptedOption": "ReturnsAccepted",
		    	"RefundOption": "MoneyBack",
		    	"ReturnsWithinOption": "Days_30",
		    	"Description": "Return policy description",
		    	"ShippingCostPaidByOption": "Buyer"
		    },
		    "ShippingDetails": {
		    	"ShippingType": "Flat",
		    	"ShippingServiceOptions": {
		    	    "ShippingServicePriority": "1",
		    	    "ShippingService": "USPSMedia",
		    	    "ShippingServiceCost": "2.50"
		    	}
		    }
		}
	    }

	    resp = self.add_item(item)

	    if resp.get('Errors'):
		print 'Item insertion failed: \n'
		print 'ErrorCode: {}'.format(resp.get('Errors')['ErrorCode'])
		print 'SeverityCode: {}'.format(resp.get('Errors')['SeverityCode'])
		print 'ErrorClassification: {} \n'.format(resp.get('Errors')['ErrorClassification'])
		print 'ErrorMessage: {} {} \n'.format(resp.get('Errors')['ShortMessage'], resp.get('Errors')['LongMessage'])
		break
	    else:
		print 'Item insertion success:'
		print 'ItemID: {}'.format(resp.get('ItemID'))
                print '------------'
		

    def add_item(self, item):

	try:
	    #api = Trading(debug=opts.debug, config_file=opts.yaml, appid=opts.appid, certid=opts.certid, devid=opts.devid, warnings=False)
	    api = Trading(domain='api.sandbox.ebay.com', config_file='api_config.yaml')
	    return api.execute('VerifyAddItem', item).dict()

	except ConnectionError as e:
	    print e
	    print e.response.dict()


def get_date():
    dt = datetime.now()
    return dt.strftime('%Y-%m-%d %H:%M:%S')

def run_main(_file):
    if not os.path.isfile(_file):
	print '{} ERROR:  {} is not a valid file.'.format(get_date(), _file) 
    ep = EbayPublisher()
    ep.process_data(_file)

if __name__ == '__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument("-file", help='The file path you want to read')
    args = parser.parse_args()

    try:
	run_main(args.file)
    except Exception as e:
    	s = str(e)
	print '{} ERROR:  {}'.format(get_date(), s)
