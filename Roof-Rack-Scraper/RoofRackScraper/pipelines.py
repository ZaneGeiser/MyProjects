# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html
from scrapy.exceptions import DropItem
from RoofRackScraper import items
import csv

class RoofrackscraperPipeline(object):
    def process_item(self, item, spider):
        if item.get('price') and item.get('title') and item.get('code'):
            item['price'] = item.get('price').strip(",$")
            return item
        else:
            raise DropItem("Missing content in %s" % item)


class CSVWriterPipeline(object):

    def open_spider(self, spider):
        self.file = open(spider.name + '_items.csv', 'w')
        self.file.write('code, price, title\n')

    def close_spider(self, spider):
        self.file.close()

    def process_item(self, item, spider):
        line = (item.get('code') + "," + item.get('price') + "," + item.get('title') + "\n")
        self.file.write(line)
        return item
