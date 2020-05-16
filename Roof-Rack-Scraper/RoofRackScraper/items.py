# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# https://docs.scrapy.org/en/latest/topics/items.html

import scrapy


class Product(scrapy.Item):
    # define the fields for your item here like:
    # name = scrapy.Field()
    title = scrapy.Field()
    code = scrapy.Field()
    price = scrapy.Field()
    on_special = scrapy.Field()
    last_updated = scrapy.Field(serializer=str)
    pass
