# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html


# useful for handling different item types with a single interface
from itemadapter import ItemAdapter
from scrapy.pipelines.files import FilesPipeline
import scrapy
from scrapy.utils.project import get_project_settings
settings = get_project_settings()


class TsinghuaPipeline:
    def process_item(self, item, spider):
        return item

class DownloadFile(FilesPipeline):
    def get_media_requests(self, item, info):
        for index, url in enumerate(item['file_urls']):
            yield scrapy.Request(url, meta={'name': item['files'][index]})

    def file_path(self, request, response=None, info=None):
        return request.meta['name']