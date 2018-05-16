import bottlenose
import xmltodict
import json
from pprint import pprint
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep

# This script pulls data from the Amazon Item Search api
#  and populates a google sheets file

# I used gspread to connect to the google sheets documentation can be found here
# https://github.com/burnash/gspread

scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]

# Sign up for google sheet api to get credentials
# https://developers.google.com/sheets/api/
creds = r""

credentials = ServiceAccountCredentials.from_json_keyfile_name(creds, scope)

gc = gspread.authorize(credentials)

# You will need to create the file on your
# Google sheets account prior to using the script

# Once created share it to the email that was
#  setup when registering for the sheets api
# The email will look similar to this:
# myprojectname@test5-194320.iam.gserviceaccount.com


# Google sheets filename
sh = gc.open('AMZ TEST')

# Worksheet name
worksheet = sh.sheet1

# Your amazon keys
AMAZON_ACCESS_KEY = ""
AMAZON_SECRET_KEY = ""
AMAZON_ASSOC_TAG = ""

amazon = bottlenose.Amazon(AMAZON_ACCESS_KEY, AMAZON_SECRET_KEY, AMAZON_ASSOC_TAG,
                           )
searches = ['Rubber Ducky']


class AMZ:
    row = 2

    def get_results(self):
        for searchterm in searches:
            results = amazon.ItemSearch(
                Keywords=searchterm,
                SearchIndex="All",
                ResponseGroup='Large,Reviews,OfferFull,Offers,OfferSummary,ItemAttributes,SalesRank')
            sleep(1)
            dictresults = dict(xmltodict.parse(results))
            for i in dictresults['ItemSearchResponse']['Items']['Item']:
                range_build = 'A' + str(self.row) + ':V' + str(self.row)
                cell_list = worksheet.range(range_build)
                cell_list[0].value = searchterm  # name
                cell_list[1].value = i.get('ItemAttributes', {}).get('Manufacturer')
                cell_list[2].value = i.get('ItemAttributes', {}).get('Title')  # name
                cell_list[3].value = i.get('ASIN')  # asin
                if i.get('BrowseNodes', {}).get('BrowseNode'):
                    for x in i.get('BrowseNodes', {}).get('BrowseNode'):
                        if not isinstance(x, str):
                            cell_list[4].value = [x.get('Ancestors', {}).get('BrowseNode', {}).get('Name'),
                                                  i.get('ItemAttributes', {}).get('Binding')][0]
                cell_list[5].value = i.get('DetailPageURL')  # url
                cell_list[6].value = i.get('OfferSummary', {}).get('TotalNew')  # Number of new for sale
                cell_list[7].value = i.get('OfferSummary', {}).get('LowestNewPrice', {}).get('FormattedPrice')
                # lowest price
                if i.get('ItemAttributes', {}).get('UPCList', {}).get('UPCListElement'):
                    cell_list[8].value = i['ItemAttributes']['UPCList']['UPCListElement'][0]  # UPC associated
                else:
                    cell_list[8].value = None
                if i.get('SalesRank'):
                    cell_list[9].value = i.get('SalesRank')  # sales rank, lower the better
                cell_list[10].value = i.get('OfferSummary', {}).get('LowestUsedPrice', {}).get('FormattedPrice')
                cell_list[11].value = i.get('OfferSummary', {}).get('TotalUsed')
                cell_list[12].value = i.get('OfferSummary', {}).get('LowestRefurbishedPrice', {}).get('FormattedPrice')
                cell_list[13].value = i['OfferSummary']['TotalRefurbished']
                cell_list[14].value = float(i.get(
                    'ItemAttributes', {}).get('PackageDimensions', {}).get('Length', {}).get(
                    '#text', '0')) * .01
                cell_list[15].value = float(i.get(
                    'ItemAttributes', {}).get('PackageDimensions', {}).get('Width', {}).get(
                    '#text', '0')) * .01
                cell_list[16].value = float(i.get(
                    'ItemAttributes', {}).get('PackageDimensions', {}).get('Height', {}).get(
                    '#text', '0')) * .01
                cell_list[17].value = float(i.get('ItemAttributes', {}).get(
                    'PackageDimensions', {}).get('Weight', {}).get(
                    '#text', '0')) * .01
                if i.get('Offers', {}).get('Offer', {}).get('OfferListing', {}).get('IsEligibleForSuperSaverShipping')\
                        == '1':
                    cell_list[18].value = "Yes"
                else:
                    cell_list[18].value = "No"
                if i.get('Offers', {}).get('Offer', {}).get('OfferListing', {}).get('IsEligibleForPrime') == '1':
                    cell_list[19].value = "Yes"
                else:
                    cell_list[19].value = "No"
                cell_list[20].value = i.get('Offers', {}).get('Offer', {}).get('Merchant', {}).get('Name')
                cell_list[21].value = i.get('CustomerReviews', {}).get('HasReviews')
                worksheet.update_cells(cell_list)
                self.row += 1


lets_go = AMZ()
lets_go.get_results()
