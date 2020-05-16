import scrapy
from RoofRackScraper import sort
from RoofRackScraper import items


class RoofRackSuperStoreSpider(scrapy.Spider):
    name = "roofracksuperstore"
    start_urls = [
    "https://www.roofracksuperstore.com.au/roof-racks/tie-down-straps/products.aspx",
    "https://www.roofracksuperstore.com.au/pioneer-products/pioneer-tradie/products.aspx",
    "https://www.roofracksuperstore.com.au/pioneer-products/pioneer-tray/products.aspx",
    "https://www.roofracksuperstore.com.au/pioneer-products/pioneer-platform/products.aspx",
    "https://www.roofracksuperstore.com.au/pioneer-products/Rhino_Rack_Backbone/products.aspx",
    "https://www.roofracksuperstore.com.au/pioneer-products/pioneer-platform-rail-kits/products.aspx",
    "https://www.roofracksuperstore.com.au/pioneer-products/pioneer-accessory-bars-rollers/products.aspx",
    "https://www.roofracksuperstore.com.au/water-sports/sup-surfboard-carriers/products.aspx",
    "https://www.roofracksuperstore.com.au/water-sports/carriers/products.aspx",
    "https://www.roofracksuperstore.com.au/water-sports/load-assist/products.aspx",
    "https://www.roofracksuperstore.com.au/water-sports/boat-loaders/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/awnings/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/shovel-spade-holder/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/recovery-track-mounts/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/jerry-can-holders/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/hi-lift-jack-brackets/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/maxtrax/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/lockn-load-platform/products.aspx",
    "https://www.roofracksuperstore.com.au/4wd/roof-top-tents/products.aspx",
    "https://www.roofracksuperstore.com.au/roof-boxes-luggage-solutions/luggage_bags/products.aspx",
    "https://www.roofracksuperstore.com.au/roof-boxes-luggage-solutions/boxes/products.aspx",
    "https://www.roofracksuperstore.com.au/roof-boxes-luggage-solutions/trays-and-baskets/products.aspx",
    "https://www.roofracksuperstore.com.au/roof-boxes-luggage-solutions/thule-luggage-and-cases/products.aspx",
    "https://www.roofracksuperstore.com.au/bike-carriers/platform-bike-carriers/products.aspx",
    "https://www.roofracksuperstore.com.au/bike-carriers/rear-mounted/products.aspx",
    "https://www.roofracksuperstore.com.au/bike-carriers/roof-mounted/products.aspx",
    "https://www.roofracksuperstore.com.au/bike-carriers/ute-indoor-mounted/products.aspx",
    "https://www.roofracksuperstore.com.au/bike-carriers/towball-carriers/products.aspx",
    "https://www.roofracksuperstore.com.au/bike-carriers/hitch-mount-carriers/products.aspx",
    "https://www.roofracksuperstore.com.au/bike-carriers/parts/products.aspx", #Spare-Wheel Mounted Bike Carriers
    "https://www.roofracksuperstore.com.au/snow-sports/carriers/products.aspx",

    ]

    def parse_product(self, response):
        def extract_with_css(query):
            return response.css(query).get(default='').strip()

        yield items.Product(
            code = sort.find_part_num(extract_with_css('h2 a::text')),
            title = extract_with_css('h2 a::text'),
            price = extract_with_css('span.price-tag::text'),
            )

    def parse(self, response):
        product_divs = response.css("div.service-wrap")
        for div in product_divs:
            yield from self.parse_product(div)
