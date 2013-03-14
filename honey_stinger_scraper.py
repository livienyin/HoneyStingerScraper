from contextlib import closing
from urllib import urlopen
import re
import xlwt

from BeautifulSoup import BeautifulSoup

def removeNonAscii(s): return "".join(i for i in s if ord(i)<128)


class HoneyStingerDealerScraper(object):

	BASE_URL = 'http://honeystinger.com/dealers.php?state={state}'

	def __init__(self):
		self.soup = None

	def open_url(self, state='OR'):
		with closing(urlopen(
			self.BASE_URL.format(
				state=state
			)
		)) as html:
			self.soup = BeautifulSoup(html.read())

	def parse_dealer(self, dealer):
		try:
			return self._parse_dealer(dealer)
		except Exception, ex:
			print ex
			raise

	def _parse_dealer(self, dealer):
		dealer_dict = {}
		try:
			dealer_dict['name'] = dealer.find(name='b').contents[0]
		except IndexError:
			return {}
		try:
			dealer_dict['phone'] = re.search(
				'T: (.*)',
				dealer.prettify()
			).group(1)
		except Exception:
			pass
		strings = [item.lstrip() for item in dealer.contents if not re.search('<br />', str(item)) and item.lstrip is not None]
		strings = [item for item in strings if item]
		strings = [item for item in strings if not re.search('T: ([0-9\-(]*)', item)]
		dealer_dict['address'] = ', '.join(strings)
		link_nodes = dealer.findAll(name='a')
		links = [link_node for link_node in link_nodes
					  if str(link_node.contents[0]) == 'website']
		dealer_dict['url'] = links[0]['href'] if links else None
		return dealer_dict

	def get_dealers(self):
		dealers = []
		table = self.soup.findAll(name='table')[0]
		rows = table.findAll(name='tr')
		for row in rows:
			dealers.extend(row.findAll(name='td'))
		return dealers

key_to_index = {
	'name': 0,
	'address': 1,
	'phone': 2,
	'url': 3
}

def write_dealer_row_to_excel(worksheet, dealer_dict, row_number):
	for key, value in dealer_dict.iteritems():
		if value:
			worksheet.write(
				row_number,
				key_to_index[key],
				label=removeNonAscii(value)
			)



	dealer_filename = 'dealers.xls'

state_list = [
	'AK',
	'AL',
	'AR',
	'AZ',
	'CA',
	'CO',
	'CT',
	'DC',
	'DE',
	'FL',
	'GA',
	'HI',
	'IA',
	'ID',
	'IL',
	'IN',
	'KS',
	'KY',
	'LA',
	'MA',
	'MD',
	'ME',
	'MI',
	'MN',
	'MO',
	'MS',
	'MT',
	'NC',
	'ND',
	'NE',
	'NH',
	'NJ',
	'NM',
	'NV',
	'NY',
	'OH',
	'OK',
	'OR',
	'PA',
	'RI',
	'SC',
	'SD',
	'TN',
	'TX',
	'UT',
	'VA',
	'VT',
	'WA',
	'WI',
	'WV',
	'WY',
]

if __name__ == '__main__':
	scraper = HoneyStingerDealerScraper()
	workbook = xlwt.Workbook()
	for state in state_list:
		scraper.open_url(state=state)
		dealers = scraper.get_dealers()
		print state

		worksheet = workbook.add_sheet(state)

		for key, column_number in key_to_index.iteritems():
			worksheet.write(
				0,
				column_number,
				key
			)

		dealers = [dealer for dealer in dealers if dealer]

		for column, dealer in enumerate(dealers):
			write_dealer_row_to_excel(
				worksheet,
				scraper.parse_dealer(dealer),
				column + 1
			)

	workbook.save('Dealers.xls')
