#!/usr/bin/env python
# coding: utf-8

# In[2]:


from docx import Document
import io
import shutil
import os
import sys
import re


# In[3]:


path = "docx/"


# In[4]:


curr_file = ""


# In[5]:


def convertDocxToText(path):
    for d in os.listdir(path):
        fileExtension=d.split(".")[-1]
        if fileExtension =="docx":
            docxFilename = path + d
            print(docxFilename)
            document = Document(docxFilename)
            textFilename = path + d.split(".")[0] + ".txt"
            global curr_file
            curr_file = textFilename
            with io.open(textFilename,"w", encoding="utf-8") as textFile:
                for para in document.paragraphs: 
                    textFile.write((para.text)+'\n')


# In[6]:


convertDocxToText(path)


# In[16]:


qbfile = open(curr_file, "r", encoding='utf-8')
paragraphs = []
ptext = ''
precinct_place_flag = 0
precinct_place_text = ''
curr_precinct_number = ''

for aline in qbfile:
    
    if precinct_place_flag == 1:
        if aline == "\n":
            precinct_place_flag = 0
            precinct_place = precinct_place_text.split('–')
            paragraphs.append('Precinct_place,'+precinct_place[1].replace(')', ''))
            precinct_place_text = ''
        else:
            precinct_place_text = precinct_place_text+aline.rstrip()+' '
            #print(precinct_place_text)
    
    if 'Количество избирательных участков, участков референдума –' in aline:
        sum_count = aline.split('–')
        sum_count[1] = re.sub('[^A-Za-z0-9]+', '', sum_count[1])
        paragraphs.append('Overall_count,'+sum_count[1])
    if 'НА ТЕРРИТОРИИ' in aline:
        township_name = aline.partition('ТЕРРИТОРИИ')[2]
        paragraphs.append('Township_name,'+township_name.strip())
    if 'Номера избирательных участков' in aline:
        precinct_nums_str = aline.partition('Номера')[2].split('-')[1].strip()
        paragraphs.append('Precinct_nums_str,'+precinct_nums_str)
        precinct_nums = precinct_nums_str.split(',')
        precinct_nums_digit = ''
        for precinct_num in precinct_nums:
            if 'по' in precinct_num or '–' in precinct_num:
                interval_parts = precinct_num.split(' ')
                begin = 0
                end = 0
                for interval_part in interval_parts:
                    interval_part = re.sub('[^A-Za-z0-9]+', '', interval_part)
                    if interval_part.isdigit() and begin == 0:
                        begin = int(interval_part.strip())
                    if interval_part.isdigit() and begin != 0:
                        end = int(interval_part.strip())
                for i in range(begin, end+1):
                    if precinct_nums_digit == '':
                        precinct_nums_digit = precinct_nums_digit+str(i)
                    else:
                        precinct_nums_digit = precinct_nums_digit+','+str(i)
            else:
                precinct_num = re.sub('[^A-Za-z0-9]+', '', precinct_num)
                if precinct_num.isdigit():
                    if precinct_nums_digit == '':
                        precinct_nums_digit = precinct_nums_digit+str(precinct_num)
                    else:
                        precinct_nums_digit = precinct_nums_digit+','+str(precinct_num)
        paragraphs.append('Precinct_nums,'+precinct_nums_digit)
                    
    if 'Количество избирательных участков, участков референдума -' in aline:
        sum_count = aline.split('-')
        sum_count[1] = re.sub('[^A-Za-z0-9]+', '', sum_count[1])
        paragraphs.append('Township_count,'+sum_count[1])
    if 'Количество избирательных участков –' in aline:
        sum_count = aline.split('–')
        sum_count[1] = re.sub('[^A-Za-z0-9]+', '', sum_count[1])
        paragraphs.append('Township_count,'+sum_count[1])        
        
    if 'Избирательный участок, участок референдума №' in aline:
        precinct_number = aline.split('№')
        precinct_number[1] = re.sub('[^A-Za-z0-9]+', '', precinct_number[1])
        paragraphs.append('Precinct_number,'+precinct_number[1])
        precinct_place_flag = 1
        curr_precinct_number = precinct_number[1]
        
    if 'Границы участка' in aline:
        if 'Улицы' in aline:
            address_text = aline.partition('Улицы:')[2].rstrip()
            streets = address_text.split(';')
            for street in streets:
                if ',' in street:
                    street_parts = street.split(',')
                        for street_part in street_parts:
                            if street_part != street_parts[0]:
                                if 'нечетная' in street_part or 'четная' in street_part:
                                    if 'нечетная' in street_part and 'четная' in street_part:
                                    if 'нечетная' in street_part:
                                    if 'четная' in street_part:
                                else:
                                paragraphs.append('Address_street,'+street_parts[0]+street_part)
                else:
                    paragraphs.append('Address_street,'+street)
        elif 'Улица' in aline:
            address_text = aline.partition('Улица')[2].rstrip()
            if ',' in address_text:
                street_parts = street.split(',')
                for street_part in street_parts:
                    if street_part != street_parts[0]:
                        paragraphs.append('Address_street,'+street_parts[0]+street_part)
        else:
            address_text = aline.split('–')[1]
            if ',' in address_text:
                street_parts = street.split(',')
                for street_part in street_parts:
                    if street_part != street_parts[0]:
                        paragraphs.append('Address_street,'+street_parts[0]+street_part)
            paragraphs.append('Address_street,'+address_text)
        
    if 'переул' in aline:
        if 'переулок' in aline:
            address_text = aline.partition('переулок')[2].rstrip()
            paragraphs.append('Address_side,'+address_text)
        elif 'переулки' in aline:
            address_text = aline.partition('переулки:')[2].rstrip()
            sides = address_text.split(';')
            for side in sides:
                paragraphs.append('Address_side,'+side)
        
    if 'террито' in aline:
        if 'территория' in aline:
            address_text = aline.partition('территория')[2].rstrip()
            paragraphs.append('Address_place,'+address_text)
        elif 'территории' in aline:
            address_text = aline.partition('территории:')[2].rstrip()
            territories = address_text.split(';')
            for territory in territories:
                paragraphs.append('Address_place,'+territory)
            
with io.open('parse_result.csv',"w", encoding="utf-8") as textFile:
    for paragraph in paragraphs: 
        textFile.write((paragraph)+'\n')
qbfile.close()


# In[ ]:




