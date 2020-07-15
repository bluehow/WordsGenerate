import re

from docx import Document
from docx.shared import Inches


def readword(filename):
    document = Document(filename)
    wordpat = r'[a-zA-Z]+'
    wordpatComplier = re.compile(wordpat)
    datasrc = ''
    for para in document.paragraphs:
        datasrc=datasrc+para.text+' '
    data=wordpatComplier.findall(datasrc)
    words =[]
    for word in data:
        if word not in words:
            words.append(word)
    return words
