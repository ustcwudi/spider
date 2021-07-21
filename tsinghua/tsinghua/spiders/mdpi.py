import scrapy
import re
import os
import xlwt
from bs4 import BeautifulSoup
from tsinghua.items import TsinghuaItem


class MdpiSpider(scrapy.Spider):
    name = 'mdpi'
    allowed_domains = ['www.mdpi.com', 'orcid.org']
    start_urls = []

    def closed(self, spider):
        self.wk.save(self.xls)
        print('爬虫结束')

    def __init__(self, p='2071-1050', v='13', n='14'):
        self.start_urls = [
            'https://www.mdpi.com/%s/%s/%s/date/default/1/1000' % (p, v, n)]
        self.xls = 'files/xls/%s-%s-%s.xls' % (p, v, n)
        self.index = 0
        self.url_list = ['']
        self.doi_list = ['']
        self.wk = xlwt.Workbook()
        self.sheet1 = self.wk.add_sheet("paper", cell_overwrite_ok=True)
        self.sheet1.write(0, 0, 'url')
        self.sheet1.write(0, 1, 'title')
        self.sheet1.write(0, 2, 'publication')
        self.sheet1.write(0, 3, 'date')
        self.sheet1.write(0, 4, 'volume')
        self.sheet1.write(0, 5, 'number')
        self.sheet1.write(0, 6, 'page')
        self.sheet1.write(0, 7, 'identifier')
        self.sheet1.write(0, 8, 'abstract')
        self.sheet1.write(0, 9, 'pdf')
        self.sheet1.write(0, 10, 'author')
        self.sheet1.write(0, 11, 'author address')
        self.sheet1.write(0, 12, 'keywords')
        self.sheet1.write(0, 13, 'cite')
        self.sheet1.write(0, 14, 'cite url')
        self.sheet1.write(0, 15, 'history')
        self.sheet1.write(0, 16, 'review')
        print('爬虫开始')

    def parse(self, response):
        html = response.text
        soup = BeautifulSoup(html, 'html5lib')
        item_list = soup.find_all('div', class_='generic-item article-item')
        for item in item_list:
            link = item.find(
                'div', class_='f-dropdown label__btn__dropdown').find('a').get('href')
            yield response.follow(link, callback=self.parse_page)

    def parse_page(self, response):
        print(response.url)
        self.index += 1
        html = response.text
        soup = BeautifulSoup(html, 'html5lib')
        self.sheet1.write(self.index, 0, response.url)
        # 标题
        title = soup.find(attrs={"name": "title"})
        self.sheet1.write(self.index, 1, title['content'])
        # prism.publicationName
        publicationName = soup.find(attrs={"name": "prism.publicationName"})
        self.sheet1.write(self.index, 2, publicationName['content'])
        # prism.publicationDate
        publicationDate = soup.find(attrs={"name": "prism.publicationDate"})
        self.sheet1.write(self.index, 3, publicationDate['content'])
        # prism.volume
        volume = soup.find(attrs={"name": "prism.volume"})
        self.sheet1.write(self.index, 4, volume['content'])
        # prism.number
        number = soup.find(attrs={"name": "prism.number"})
        self.sheet1.write(self.index, 5, number['content'])
        # prism.startingPage
        startingPage = soup.find(attrs={"name": "prism.startingPage"})
        self.sheet1.write(self.index, 6, startingPage['content'])
        # identifier
        identifier = soup.find(attrs={"name": "dc.identifier"})
        self.sheet1.write(self.index, 7, identifier['content'])
        # 摘要
        abstract = soup.find(attrs={"name": "description"})
        if len(abstract['content']) > 30000:
            self.sheet1.write(self.index, 8, abstract['content'][0:30000])
        else:
            self.sheet1.write(self.index, 8, abstract['content'])
        # pdf
        pdf = soup.find(attrs={"name": "fulltext_pdf"})
        self.sheet1.write(self.index, 9, pdf['content'])
        # 作者
        author1 = ''
        author2 = ''
        author_list = soup.find(
            'div', class_='art-authors hypothesis_container').find_all('div', class_="sciprofiles-link")
        for author in author_list:
            a = author.find('a')
            a_href = ''
            if a is not None:
                a_href = a.get('href')
            orc = ''
            links = author.parent.find_all('a')
            for link in links:
                if link.get('href').startswith('https://orcid.org/'):
                    orc = link.get('href')[18:]
                    item = TsinghuaItem()
                    files = []
                    file_urls = []
                    if not os.path.exists('files/orc/'+orc+'.person.json'):
                        files.append('orc/'+orc+'.person.json')
                        file_urls.append(link.get('href')+'/person.json')
                    else:
                        print('exist ' + orc+'.person.json')
                    if not os.path.exists('files/orc/'+orc+'.affiliation.json'):
                        files.append('orc/'+orc+'.affiliation.json')
                        file_urls.append(link.get('href') +
                                         '/affiliationGroups.json')
                    if not os.path.exists('files/orc/'+orc+'.works.json'):
                        files.append('orc/'+orc+'.works.json')
                        file_urls.append(link.get(
                            'href')+'/worksPage.json?offset=0&sort=date&sortAsc=false&pageSize=100')
                    item['files'] = files
                    item['file_urls'] = file_urls
                    yield item
            author1 += author.get_text()+'|' + author.parent.find('sup').get_text().strip() + \
                '|' + orc+'|' + a_href+'\n'
        self.sheet1.write(self.index, 10, author1)
        # 作者地址
        author_address_list = soup.find_all('div', class_='affiliation')
        for author_address in author_address_list:
            key = author_address.find('div', class_='affiliation-item')
            if key is not None:
                author2 += key.get_text() + ':'+author_address.find('div',
                                                                    class_='affiliation-name').get_text() + '\n'
            else:
                author2 += author_address.find('div',
                                               class_='affiliation-name').get_text() + '\n'
        self.sheet1.write(self.index, 11, author2)
        # 关键词
        keyword_array = ''
        keywords = soup.find(
            'div', class_='art-keywords in-tab hypothesis_container')
        if keywords is not None:
            for keyword in keywords.find_all('a'):
                keyword_array += keyword.get_text() + '\n'
            self.sheet1.write(self.index, 12, keyword_array)
        # history
        history = soup.find('div', class_='pubhistory')
        self.sheet1.write(self.index, 15, history.get_text())
        # add list
        self.url_list.append(response.url)
        self.doi_list.append(identifier['content'])
        # 引用
        match = re.search(r'\"/cite-count/(.*?)\"', html, re.I)
        if match:
            yield response.follow('/cite-count/'+match.group(1), callback=self.parse_cite)
        # review
        buttons = soup.find_all('a', class_='button button--color-inversed')
        for button in buttons:
            if(button is not None and button.get('href') is not None and button.get('href').endswith('review_report')):
                yield response.follow(button['href'], callback=self.parse_review)
        pass

    def parse_review(self, response):
        item = TsinghuaItem()
        file_name = []
        file_url = []
        i = self.url_list.index(response.url[0:-14])
        print(i)
        html = response.text
        soup = BeautifulSoup(html, 'html5lib')
        content = soup.find('div', class_='abstract_div').contents[-2]
        list = []
        words = ''
        for p in content.contents:
            if p.name is not None:
                if re.search(r'Reviewer [0-9] Report', p.get_text(), re.I) or re.search(r'Round [0-9]', p.get_text(), re.I) or p.get_text() == ("Author Response"):
                    list.append(words)
                    words = p.get_text()+'\n'
                else:
                    words = (words+p.get_text()+'\n')
                    attachments = p.find_all('a')
                    for attachment in attachments:
                        href = attachment['href']
                        result = re.search(
                            r'/([0-9]+)/(.*)\?file=(.*)&report=([0-9]+)', href, re.I)
                        if result:
                            name = 'response/'+result.group(
                                1)+'.'+result.group(2)+'.'+result.group(3)+'.'+result.group(4)+os.path.splitext(attachment.get_text())[-1]
                            if not os.path.exists('files/'+name):
                                print('need ' + name)
                                file_name.append(name)
                                file_url.append(href)
                            else:
                                print('exist ' + name)
                            words = words+name+'\n'
                        else:
                            print(href)
        list.append(words)

        index = 0
        for line in list:
            if len(line) > 30000:
                self.sheet1.write(i, 16+index, line[0:30000])
            else:
                self.sheet1.write(i, 16+index, line)
            index += 1
        # reviewer
        reviewer_list = soup.find_all(
            'div', style='display: block;font-size:14px; line-height:30px;')
        reviewers = ''
        for reviewer in reviewer_list:
            reviewers += reviewer.get_text() + '\n'
        self.sheet1.write(i, 16, reviewers)

        item['files'] = file_name
        item['file_urls'] = file_url
        yield item

    def parse_cite(self, response):
        index = self.doi_list.index(response.url[32:].replace('%252F', '/'))
        print(index)
        html = response.text
        soup = BeautifulSoup(html, 'html5lib')
        cites = soup.find_all('div', class_='relative-size-container')
        cite1 = ''
        cite2 = ''
        for cite in cites:
            title = cite.find('div', class_='relative-size-title')
            count = cite.find('a')
            cite1 += title.get_text().strip() + ':'+count.get_text().strip() + '\n'
            cite2 += count.get('href').strip() + '\n'
        self.sheet1.write(index, 13, cite1)
        self.sheet1.write(index, 14, cite2)
