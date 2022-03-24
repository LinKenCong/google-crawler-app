import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse
import time

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36'}


def search_google_result(keyword):
    num = 0
    searchResultAll = []
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36'}

    def search(num):
        html = requests.get(
            'https://www.google.com/search?q={}&btnG=Search&gbv=10&start={}'.format(keyword, num), headers=headers, timeout=30).text
        soup = BeautifulSoup(html, 'html.parser')
        searchResult = list(soup.find_all('div', class_="tF2Cxc"))
        return searchResult

    while(len(searchResultAll) <= 20):
        searchResultAll += search(num)
        num += 10
        time.sleep(5)
    return searchResultAll[0:20]


def get_kws_research(searchResult, keyword_val):
    title = []
    link = []
    domain = []
    description = []
    keyword = []
    for i in searchResult:
        title_val = i.find('h3')
        link_val = i.find('a', href=True)['href']
        domain_val = urlparse(link_val).netloc
        description_val = i.find(
            'div', class_="VwiC3b yXK7lf MUxGbd yDYNvb lyLwlc lEBKkf")
        if(title_val and link_val):
            keyword.append(keyword_val)
            title.append(title_val.text)
            link.append(link_val.split('&')[0])
            domain.append(domain_val)
            if(description_val):
                description.append(description_val.text)
            else:
                description.append('no description')
    result = [keyword, title, link, domain, description]
    return result


def get_kws_research_more(arrayData, searchFunc):
    result = []
    for keyword in arrayData:
        val = searchFunc(keyword)
        try:
            searchResult = get_kws_research(val, keyword)
            time.sleep(15)
        except:
            searchResult = val
        result.append(searchResult)
    return result
