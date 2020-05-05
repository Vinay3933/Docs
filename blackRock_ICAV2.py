# -*- coding: utf-8 -*-
"""
Created on Fri Apr 24 17:33:08 2020

@author: c-VrushabhMa
"""

import re
import os
import time
import pandas as pd
from lxml import etree
from PyPDF2 import PdfFileReader
import xml.etree.ElementTree as ET

def to_xml(file_path):
    poppler_path = r'C:\Users\C-vrushabhma\Anaconda\pkgs\poppler-0.87.0-h0cd1227_1\Library\bin\pdftohtml.exe'
    
    cmd = poppler_path+" -i -c -hidden -fontfullname -xml "+file_path+" temp.xml"
    
    os.popen(cmd)
    time.sleep(5)

    while True:
        if os.path.isfile('temp.xml') and os.stat('temp.xml').st_size/1024 > 100:
                break
        time.sleep(0.5)
    
    xml_file = "temp.xml"
    parser = etree.XMLParser(recover=True)
    tree = ET.parse(xml_file,parser=parser)
        
    root = tree.getroot() 
    
    return root


def get_text(root,page_range):
    #page = root.xpath('//page[@number="87"]')[0]
    pages = root.xpath('//page')
    main_text = ""
    
    for page in pages[page_range[0]:page_range[1]]:
        text_tags = page.xpath('text')
        prev_top = text_tags[0].attrib['top']
        t = ''.join(text_tags[0].itertext())
        for tag in text_tags[1:]:
            diff = abs(int(tag.attrib['top'])-int(prev_top))
            if diff > 25:
                t = t+'\n\n'+''.join(tag.itertext())#.strip()
            elif diff <= 25 and diff > 5:
                t = t+'\n'+''.join(tag.itertext())#.strip()
            else:
                t = t+' '.join(tag.itertext())#.strip()
            #t = t+temp_text
            prev_top = tag.attrib['top']               
        #text = ''.join(page.itertext())
        if main_text:
            main_text = main_text+'\n'+t
        else:
            main_text = t
    
    return main_text


def get_page_range(root,fund_list,key_words):
    pages = []
    patt = '|'.join(key_words)
    
    for fund_name in fund_list:
        print(fund_name)
        fund_name = fund_name.replace("*","")
        #xpath = '//b[contains(text(),"'+fund_name+'")]'
        xpath = '//b[text()="'+fund_name+'" or text()="'+fund_name+' " or contains(text(),"'+fund_name+' ")]'
        ele = root.xpath(xpath)
        page_path = 'parent::text/parent::page[1]'
        if not ele:
            font_values =[21,22,23]
            x_p = ' or '.join(['@font="'+str(i)+'"' for i in font_values])
            xpath = '//text[(text()="'+fund_name+'" or text()="'+fund_name+' " or contains(text(),"'+fund_name+' ")) and '+'('+x_p+')]'
            ele = root.xpath(xpath)
            page_path = 'parent::page[1]'
        
        if ele:
            for i in ele:
                page = i.xpath(page_path)
                if page:
                    page = page[0]
                    s = ''.join(page.itertext())
                    if re.search(patt, s):            
                        page_number = page.attrib['number']
                        print(page_number)
                        pages.append(int(page_number))
                #pages[int(page_number)] = fund_name
    
    #pages = sorted(pages)
    
    return min(pages),max(pages)


def get_pattern(fund_list,last_word):
    patt_list = []
    for ind,i in enumerate(fund_list):
        
        j = '|'.join(['(?:^ |^)'+j.replace('*','').replace('(','\(').replace(')','\)') for j in fund_list if j != i])
        j = j+'|^'+last_word
        i = i.replace('*','').replace('(','\(').replace(')','\)')
        pattern  = '(?=(?:^ |^)'+i+'(?:|\s+)$)[^$$]+?(?='+j+')'    
        patt_list.append(pattern)

    return patt_list
 


def get_data(sub_text):
    asset_manager = 'BlackRock Global Funds'
    fund_name = ''
    inv_obj = ''
    strategy = ''
    regulation = ''
    derivative = ''
    leverage = ''
    
    patt = '^.+'
    temp = re.findall(patt, sub_text)
    if temp:
        fund_name = temp[0].replace('-','')
        fund_name = re.sub('\n+',' ',fund_name.strip())
    
    #patt = '(?<=Investment Objective)[^$$]+?(?<=\n\n)'
    """Pattern Explaination

       (?<=^Investment Objective) - Start matching from this     
       [^$$]+?                    - match until $$ found
       (?=Investment Policy)      - Matches upto Investment policy
       
       Note
       (?<=word) - puts cursor at end of word
       (?=word) - cursror at star of the word

    """
    patt = '(?<=^Investment Objective)[^$$]+?(?=Investment Policy)'
    
    temp = re.findall(patt, sub_text,re.MULTILINE)
    if temp:
        inv_obj = temp[0].strip().replace('\n',' ')
    
    patt = '(?<=Investment Policy)[^$$]+?(?=In implementing its investment policy)'    
    temp = re.findall(patt, sub_text,re.MULTILINE)
    if temp:
        derivative = re.sub('\n+',' ',temp[0].strip())
    
    patt = '(?=In implementing its investment policy)[^$$]+?(?:(?<=derivative instruments\.)|(?=\n\s))'
    
    #patt = '(?=In implementing its investment (?:policy|))[^$$]+?(?=\n\s)'
    temp = re.findall(patt, sub_text,re.MULTILINE)
    if temp:
        leverage = re.sub('\n+',' ',temp[0].strip())
    
    return [asset_manager,fund_name,inv_obj,strategy,regulation,derivative,leverage]

file_path = 'prospectus--blackrock-funds-i-icav-23-dec-2019-emea.pdf'
pdf_link = 'https://www.blackrock.com/ch/individual/en/literature/prospectus/prospectus--blackrock-funds-i-icav-23-dec-2019-emea.pdf'

xml_doc = to_xml(file_path)

text = get_text(xml_doc, [3,4])
fund_list = re.findall('^BlackRock.+(?=\n\n)',text,re.M)
fund_list = [i.strip() for i in fund_list]

key_words = ['The investment objective']
start_pg,end_pg = get_page_range(xml_doc, fund_list,key_words)

text = get_text(xml_doc,[start_pg-2,end_pg+3])

last_word = 'APPENDIX B'

patt_list = get_pattern(fund_list,last_word)

sub_text_list = []
for ind,patt in enumerate(patt_list):
    i = re.findall(patt,text,re.M)
    if i:
        sub_text_list.append(i[0])
    else:
        print(ind)

sub_text_list = [re.sub('[ ]{2}',' ',i) for i in sub_text_list]

data =[]
for sub_text in sub_text_list:
    d = get_data(sub_text)
    data.append(d)
    print(d[1],"Done")

text = get_text(xml_doc,[0,1])
regulation = re.findall('(?=\(an Irish)[^$$]+?(?=\()',text,re.M)
if regulation:
    regulation = regulation[0].replace('(','').strip()

columns = ['Asset Manager', 'Fund Name', 'Investment_Objective', 'Strategy',
            'Regualtion_Text', 'Derivative_Text', 'Leverage_Text']

df = pd.DataFrame(data=data,columns = columns)

df['Regualtion_Text'] = regulation
df['pdf_link'] = pdf_link

df = df.applymap(lambda x: 'NA' if not x else x)
df.to_excel('BlackrockICAVFunds.xlsx',index=False)
