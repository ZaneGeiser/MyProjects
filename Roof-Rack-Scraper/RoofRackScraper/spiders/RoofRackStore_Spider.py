import scrapy
from RoofRackScraper import sort
from RoofRackScraper import items


class RoofRackStoreSpider(scrapy.Spider):
    name = "roofrackstore"
    handle_httpstatus_list = [501]
    start_urls = [
        "https://www.roofrackstore.com.au/roof-rack-bar-packs-c-43.html",
        "https://www.roofrackstore.com.au/roof-rack-foot-packs-c-42.html",
        "https://www.roofrackstore.com.au/indoor-ute-bicycle-carriers-c-2_29.html",
        "https://www.roofrackstore.com.au/hitch-bicycle-carriers-c-2_28_32.html",
        "https://www.roofrackstore.com.au/spare-wheel-bicycle-carriers-c-2_28_34.html",
        "https://www.roofrackstore.com.au/tow-ball-bicycle-carriers-c-2_28_31.html",
        "https://www.roofrackstore.com.au/roof-mounted-bicycle-carriers-c-2_27.html",
        "https://www.roofrackstore.com.au/bike-carrier-accessories-c-2_118.html",
        "https://www.roofrackstore.com.au/awnings-for-vehicles-c-94_112.html",
        "https://www.roofrackstore.com.au/camping-accessories-c-94_114.html",
        "https://www.roofrackstore.com.au/platform-accessories-c-94_115.html",
        "https://www.roofrackstore.com.au/platform-racks-c-94_111.html",
        "https://www.roofrackstore.com.au/roof-baskets-and-accessories-c-94_110.html",
        "https://www.roofrackstore.com.au/roof-top-tents-c-94_113.html",
        "https://www.roofrackstore.com.au/trade-accessories-c-94_116.html",
        #"https://www.roofrackstore.com.au/rhino-rack-roof-rack-fitting-kits-c-44_109.html",
        #"https://www.roofrackstore.com.au/thule-roof-rack-fitting-kits-c-44_96.html",
        #"https://www.roofrackstore.com.au/yakima-whispbar-roof-rack-fitting-kits-c-44_97.html",
        "https://www.roofrackstore.com.au/roof-boxes-c-3.html",
        "https://www.roofrackstore.com.au/snowsports-c-22.html",
        "https://www.roofrackstore.com.au/thule-chariot-c-100_101.html",
        "https://www.roofrackstore.com.au/thule-chariot-accessories-c-100_103.html",
        "https://www.roofrackstore.com.au/thule-chariot-conversion-kits-c-100_102.html",
        "https://www.roofrackstore.com.au/surfboard-covers-c-21_117.html",
        "https://www.roofrackstore.com.au/roof-mounted-kayak-carriers-c-21_86.html",
        "https://www.roofrackstore.com.au/surfboard-sup-carriers-c-21_85.html",
        "https://www.roofrackstore.com.au/tie-downs-accessories-c-21_88.html",
        "https://www.roofrackstore.com.au/ute-hitch-carriers-c-21_87.html",
        "https://www.roofrackstore.com.au/luggage-cases-travel-bags-c-92.html",
    ]

    def parse_product(self, response):
        def extract_with_css(query):
            return response.css(query).get(default='').strip()

        yield items.Product(
            code = sort.find_part_num(extract_with_css('a.product-info-text::text')),
            title = extract_with_css('a.product-info-text::text'),
            price = extract_with_css('span.priceNormal::text'),
            )

    def parse(self, response):
        product_divs = response.css("div.product-item")
        for div in product_divs:
            yield from self.parse_product(div)

        next_page = response.css('a[title=" Next Page "]::attr(href)').get()
        if next_page is not None:
            yield response.follow(next_page, callback=self.parse)
