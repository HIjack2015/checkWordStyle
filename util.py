from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from loguru import logger

from str_util import find

def export_xml(document):
    body_element = document._body._body
    a=open("pic_align.xml",encoding="utf-8",mode="a")
    a.write(body_element.xml)
    a.close()



export_xml(Document("pic_align.docx"))