import urllib3
from bs4 import BeautifulSoup as bs4
from docx import Document
from docx.shared import Inches

def getdictforword(word):
    url ='http://www.iciba.com/word?w='
    url_word = url+word
    http = urllib3.PoolManager()
    html = http.request('GET',url_word).data.decode('utf-8')
    word_dict={}
    soup = bs4(html,'html.parser')
    word_dict['发音']=soup.select('.Mean_symbols__5dQX7')[0].text
    word_dict['释义']=soup.select('.Mean_part__1RA2V')[0].text
    try:
        word_dict['例句1'] = soup.select('.NormalSentence_sentence__3q5Wk')[0].text
    except:
        word_dict['例句1'] = ''
    try:
        word_dict['例句2']=soup.select('.NormalSentence_sentence__3q5Wk')[1].text
    except:
        word_dict['例句2'] = ''
    print(word,'done!')
    print(word_dict)
    return word_dict


