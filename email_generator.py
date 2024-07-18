#scope: The goal for this project is to directly create email to send to Sarah with existing excel sheet

import csv
import sys

class CSVParser:
    
    def __init__(self, filename):
        self.csv_lines = []
        with open(filename, newline='', encoding='utf8') as csvfile:
            sheet_reader = csv.reader(csvfile, delimiter=' ', quotechar='|')
            for row in sheet_reader:
                line = ','.join(row)
                if line.startswith('\ufeff'):
                    self.csv_lines.append(line[1:])
                else:
                    self.csv_lines.append(line)
                
    def chart_split(self):
        content = []
        self.funding_source = []
        self.purchase_item = []
        for line in self.csv_lines:
            items = line.split(',')
            while '' in items:
                items.remove('')
            content.append(items)
        first_sheet_index = content.index([])
        second_sheet_index = 0
        for i, val in enumerate(content): 
            if val == []:
                second_sheet_index = i
        self.funding_source = content[0: first_sheet_index]
        self.purchase_item = content[second_sheet_index+1:]

    def process_funding_source(self):
        START = 1
        funding_source = ''
        for index, item in enumerate(self.funding_source[START:], START):
            funding_source += item[0] + '(' + item[1] + ')' + ' - ' + item[2] + '\n'
        return funding_source
    
    def process_purchase_item(self):
        START = 1
        purchase_item = ''
        for index, item in enumerate(self.purchase_item[START:], START):
            purchase_item += 'item' + item[0] + ':' + '\n'
            purchase_item += '  - Name: ' + item[1] + '\n'
            purchase_item += '  - link: ' + item[2] + '\n'
            purchase_item += '  - Item Number: ' + item[5] + '\n'
            purchase_item += '  - price: ' + item[3] + '\n'
            purchase_item += '  - description: ' + item[4] + '\n'
        return purchase_item


if __name__ == '__main__':
    if sys.argv[1].endswith('.csv'):
        file_handler = CSVParser(sys.argv[1].strip())
        file_handler.chart_split()
        funding = file_handler.process_funding_source()
        print(funding)
        purchase = file_handler.process_purchase_item()
        print(purchase)
        