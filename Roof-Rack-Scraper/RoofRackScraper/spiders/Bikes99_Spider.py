import scrapy
from RoofRackScraper import items


class Bikes99Spider(scrapy.Spider):
    name = "bikes99"
    start_urls = [
        "https://www.99bikes.com.au/car-racks"
    ]

    def parse_product(self, response):
        def extract_with_css(query):
            return response.css(query).get(default='').strip()

        yield items.Product(
            code = extract_with_css('a::attr(data-id)'),
            title = extract_with_css('a::attr(data-name)'),
            price = extract_with_css('a::attr(data-price)'),
            )

    def parse(self, response):
        product_anchors = response.css('a.product-item-photo')
        for anchor in product_anchors:
            yield from self.parse_product(anchor)

        next_page = response.css('.next::attr(href)').get()
        if next_page is not None:
            yield response.follow(next_page, callback=self.parse)
