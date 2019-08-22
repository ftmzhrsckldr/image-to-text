# -*- coding: utf-8 -*-
"""
Created on Thu May  3 15:32:25 2018
@author: 672028
"""

from docx import Document
from docx.shared import Inches
from docx.shared import Cm


def Documentmaker(img_path, doc_path):
    document = Document()
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(0.17)
        section.bottom_margin = Cm(0.17)
        section.left_margin = Cm(0.17)
        section.right_margin = Cm(0.17)
    document.add_picture(path, width=Inches(7.38), height=Inches(10.77))
    document.save(doc_path)