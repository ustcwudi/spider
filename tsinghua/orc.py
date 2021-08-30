import xlrd
import json
import os
import sys
import urllib.request
from xlutils.copy import copy


def time_str(year, month, day):
    year_str = ''
    if year is not None:
        year_str = year
    else:
        year_str = '0000'
    month_str = ''
    if month is not None:
        month_str = month
    else:
        month_str = '00'
    day_str = ''
    if day is not None:
        day_str = day
    else:
        day_str = '00'
    return year_str+month_str+day_str


def null_str(content):
    if content is None:
        return ''
    else:
        return content


def orc(file):
    readbook = xlrd.open_workbook(file)
    writebook = copy(readbook)
    if len(readbook.sheets()) == 1:
        writebook.add_sheet('orc')
    writesheet = writebook.get_sheet(1)
    writesheet.write(0, 0, 'orc')
    writesheet.write(0, 1, 'title')
    writesheet.write(0, 2, 'name')
    writesheet.write(0, 3, 'country')
    writesheet.write(0, 4, 'keyword')
    writesheet.write(0, 5, 'identifier')
    writesheet.write(0, 6, 'employment')
    writesheet.write(0, 7, 'qualification')
    writesheet.write(0, 8, 'education')
    writesheet.write(0, 9, 'works')
    writesheet.write(0, 10, 'reviews')
    sheet = readbook.sheet_by_index(0)
    list = []
    nrows = sheet.nrows
    for row in range(1, nrows):
        authors = sheet.cell(row, 10).value
        author_list = authors.split('\n')
        for author in author_list:
            if len(author) > 1:
                spans = author.split('|')
                if len(spans[2]) == 19:
                    if not len(spans[2]) in list:
                        list.append(spans[2])
    i = 0
    opener = urllib.request.build_opener()
    opener.addheaders = [
        ('User-agent', 'Opera/9.80 (Android 2.3.4; Linux; Opera Mobi/build-1107180945; U; en-GB) Presto/2.8.149 Version/11.10')]
    urllib.request.install_opener(opener)
    for orc in list:
        print(orc)
        i += 1
        if not os.path.exists('files/orc/'+orc+".review.json"):
            urllib.request.urlretrieve('https://orcid.org/'+orc +
                                       '/peer-reviews.json?offset=&sortAsc=false', 'files/orc/'+orc+".review.json")
        if not os.path.exists('files/orc/'+orc+".person.json"):
            urllib.request.urlretrieve('https://orcid.org/'+orc +
                                       '/person.json', 'files/orc/'+orc+".person.json")
        if not os.path.exists('files/orc/'+orc+".affiliation.json"):
            urllib.request.urlretrieve('https://orcid.org/'+orc +
                                       '/affiliationGroups.json', 'files/orc/'+orc+".affiliation.json")
        if not os.path.exists('files/orc/'+orc+".works.json"):
            urllib.request.urlretrieve('https://orcid.org/'+orc +
                                       '/worksPage.json?offset=0&sort=date&sortAsc=false&pageSize=100', 'files/orc/'+orc+".works.json")
        with open("files/orc/"+orc+".person.json", "r", encoding='utf-8') as f:
            dict = json.load(f)
            # 0 id
            writesheet.write(i, 0, orc)
            # 1 title
            writesheet.write(i, 1, dict['title'])
            # 2 name
            writesheet.write(i, 2, dict['displayName'])
            # 3 country
            content = ''
            if dict['countryNames'] is not None:
                for key in dict['countryNames']:
                    content += (dict['countryNames'][key]+'\n')
            writesheet.write(i, 3, content)
            # 4 keyword
            content = ''
            if dict['publicGroupedKeywords'] is not None:
                for key in dict['publicGroupedKeywords']:
                    content += (key+'\n')
            writesheet.write(i, 4, content)
            # 5 identifier
            content = ''
            if dict['publicGroupedPersonExternalIdentifiers'] is not None:
                for key in dict['publicGroupedPersonExternalIdentifiers']:
                    content += key
                    content_ = ''
                    for id in dict['publicGroupedPersonExternalIdentifiers'][key]:
                        if id['url'] is not None:
                            content_ += (null_str(id['url']['value'])+'|')
                    content += ':'+content_+'\n'
            writesheet.write(i, 5, content)
        with open("files/orc/"+orc+".affiliation.json", "r", encoding='utf-8') as f:
            dict = json.load(f)
            # 6 employment
            content = ''
            for employment in dict['affiliationGroups']['EMPLOYMENT']:
                content += (null_str(employment['defaultAffiliation']['affiliationName']['value'])+'|'
                            + null_str(employment['defaultAffiliation']['city']['value'])+'|'
                            + null_str(employment['defaultAffiliation']['region']['value'])+'|'
                            + null_str(employment['defaultAffiliation']['country']['value'])+'|'
                            + null_str(employment['defaultAffiliation']['departmentName']['value'])+'|'
                            + null_str(employment['defaultAffiliation']['roleTitle']['value'])+'|'
                            + null_str(employment['defaultAffiliation']['affiliationType']['value'])+'|'
                            + time_str(employment['defaultAffiliation']['startDate']['year'], employment['defaultAffiliation']
                                       ['startDate']['month'], employment['defaultAffiliation']['startDate']['day'])+'|'
                            + time_str(employment['defaultAffiliation']['endDate']['year'], employment['defaultAffiliation']
                                       ['endDate']['month'], employment['defaultAffiliation']['endDate']['day'])
                            + '\n')
            if len(content) > 30000:
                content = content[:30000]
            writesheet.write(i, 6, content)
            # 7 qualification
            content = ''
            for qualification in dict['affiliationGroups']['QUALIFICATION']:
                content += (
                    null_str(qualification['defaultAffiliation']
                             ['affiliationName']['value'])+'|'
                    + null_str(qualification['defaultAffiliation']['city']['value'])+'|'
                    + null_str(qualification['defaultAffiliation']['region']['value'])+'|'
                    + null_str(qualification['defaultAffiliation']['country']['value'])+'|'
                    + null_str(qualification['defaultAffiliation']['roleTitle']['value'])+'|'
                    + null_str(qualification['defaultAffiliation']
                               ['departmentName']['value'])+'|'
                    + null_str(qualification['defaultAffiliation']
                               ['affiliationType']['value'])+'|'
                    + time_str(qualification['defaultAffiliation']['startDate']['year'], qualification['defaultAffiliation']
                               ['startDate']['month'], qualification['defaultAffiliation']['startDate']['day'])+'|'
                    + time_str(qualification['defaultAffiliation']['endDate']['year'], qualification['defaultAffiliation']
                               ['endDate']['month'], qualification['defaultAffiliation']['endDate']['day'])
                    + '\n')
            if len(content) > 30000:
                content = content[:30000]
            writesheet.write(i, 7, content)
            # 8 education
            content = ''
            for education in dict['affiliationGroups']['EDUCATION']:
                content += (
                    null_str(education['defaultAffiliation']
                             ['affiliationName']['value'])+'|'
                    + null_str(education['defaultAffiliation']
                               ['city']['value'])+'|'
                    + null_str(education['defaultAffiliation']
                               ['region']['value'])+'|'
                    + null_str(education['defaultAffiliation']
                               ['country']['value'])+'|'
                    + null_str(education['defaultAffiliation']
                               ['roleTitle']['value'])+'|'
                    + null_str(education['defaultAffiliation']
                               ['departmentName']['value'])+'|'
                    + null_str(education['defaultAffiliation']
                               ['affiliationType']['value'])+'|'
                    + time_str(education['defaultAffiliation']['startDate']['year'], education['defaultAffiliation']
                               ['startDate']['month'], education['defaultAffiliation']['startDate']['day'])+'|'
                    + time_str(education['defaultAffiliation']['endDate']['year'], education['defaultAffiliation']
                               ['endDate']['month'], education['defaultAffiliation']['endDate']['day'])
                    + '\n')
            if len(content) > 30000:
                content = content[:30000]
            writesheet.write(i, 8, content)
        with open("files/orc/"+orc+".works.json", "r", encoding='utf-8') as f:
            dict = json.load(f)
            # 9 works
            content = ''
            for group in dict['groups']:
                for work in group['works']:
                    content += (
                        null_str(work['sourceName'])+'|'
                        + null_str(work['title']['value'])+'|'
                        + null_str(work['workType']['value']))+'|'
                    if work['publicationDate'] is not None:
                        content += time_str(work['publicationDate']['year'],
                                            work['publicationDate']['month'], work['publicationDate']['day'])+'|'
                    if work['workExternalIdentifiers'] is not None:
                        for id in work['workExternalIdentifiers']:
                            content += null_str(id['externalIdentifierType']['value']) + \
                                ':' + \
                                null_str(id['externalIdentifierId']
                                         ['value'])+'|'
                    content += '\n'
            if len(content) > 30000:
                content = content[:30000]
            writesheet.write(i, 9, content)
        with open("files/orc/"+orc+".review.json", "r", encoding='utf-8') as f:
            dict = json.load(f)
            # 10 reviews
            j = 0
            for review in dict:
                content = null_str(review['name'])+'|'+null_str(review['type'])+'|' + \
                    null_str(review['groupType'])+'|' + \
                    null_str(review['groupIdValue'])+'\n'
                for group in review['peerReviewDuplicateGroups']:
                    for peer in group['peerReviews']:
                        content += (null_str(peer['role']['value'])+'|'+null_str(peer['type']['value'])+'|'+null_str(peer['orgName']['value'])+'|'+time_str(
                            peer['completionDate']['year'], peer['completionDate']['month'], peer['completionDate']['day'])+'\n')
                if len(content) > 30000:
                    content = content[:30000]
                writesheet.write(i, 10+j, content)
                j = j+1
                if j > 240:
                    break
    writebook.save(file)


if __name__ == '__main__':
    arg1 = sys.argv[1]
    arg2 = sys.argv[2]
    arg3 = sys.argv[3]
    orc('files/xls/'+arg1+'-'+arg2+'-'+arg3+'.xls')
