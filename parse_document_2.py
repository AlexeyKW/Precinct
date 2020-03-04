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


# In[7]:


def object_type(p_type):
    object_string = 'none'
    if 'улицы' == p_type:
        object_string = 'Address_street, улица'
    elif 'Улицы' == p_type:
        object_string = 'Address_street, улица'
    elif 'Улица' == p_type:
        object_string = 'Address_street, улица'
    elif '– Проспект' == p_type:
        object_string = 'Address_street, проспект'
    elif 'переулок' == p_type:
        object_string = 'Address_street, переулок'
    elif 'переулки' == p_type:
        object_string = 'Address_street, переулок'
    elif 'территория' == p_type:
        object_string = 'Address_street,'
    elif 'территории' == p_type:
        object_string = 'Address_street,'
    elif '– Площадь' == p_type:
        object_string = 'Address_street, площадь'
    else:
        object_string = 'none'
    return object_string


# In[33]:


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
            paragraphs.append('Precinct_place,'+precinct_place[1].replace(')', '').replace(',', ''))
            #print (precinct_place[1])
            if '»,' in precinct_place[1]:
                address_parts = precinct_place[1].replace(')', '').rpartition('»,')
                num = len(address_parts)
                #print(num)
                paragraphs.append('Precinct_address,'+address_parts[num-1].replace(',', ''))
            if '),' in precinct_place[1]:
                paragraphs.append('Precinct_address,'+precinct_place[1].partition('),')[2].replace(')', '').replace(',', ''))
            if '»,' not in precinct_place[1] and '),' not in precinct_place[1]:
                precinct_address_parts = precinct_place[1].split(',')
                num = len(precinct_address_parts)
                paragraphs.append('Precinct_address,'+precinct_address_parts[num-2]+' '+precinct_address_parts[num-1].replace(')', ''))
            precinct_place_text = ''
        else:
            precinct_place_text = precinct_place_text+aline.rstrip()+' '
            
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
                    
    if 'Количество избирательных участков, участков референдума -' in aline or 'Количество избирательных участков –' in aline:
        if '-' in aline:
            sum_count = aline.split('-')
        if '–' in aline:
            sum_count = aline.split('–')
        sum_count[1] = re.sub('[^A-Za-z0-9]+', '', sum_count[1])
        paragraphs.append('Township_count,'+sum_count[1])
        
    if 'Избирательный участок, участок референдума №' in aline:
        precinct_number = aline.split('№')
        precinct_number[1] = re.sub('[^A-Za-z0-9]+', '', precinct_number[1])
        paragraphs.append('Precinct_number,'+precinct_number[1])
        precinct_place_flag = 1
        curr_precinct_number = precinct_number[1]       
    
    parse_type = 'none'
    if 'улицы' in aline:
        parse_type = 'улицы'
    elif 'Улицы' in aline:
        parse_type = 'Улицы'
    elif 'Улица' in aline:
        parse_type = 'Улица'
    elif '– Проспект' in aline:
        parse_type = '– Проспект'
    elif 'переулок' in aline:
        parse_type = 'переулок'
    elif 'переулки' in aline:
        parse_type = 'переулки'
    elif 'территория' in aline:
        parse_type = 'территория'
    elif 'территории' in aline:
        parse_type = 'территории'
    elif '– Площадь' in aline:
        parse_type = '– Площадь'
    else:
        parse_type = 'none'
    
    if parse_type != 'none':
        address_text = aline.partition(parse_type)[2].rstrip()
        #print (address_text)
        if address_text != '':
            #print(address_text)
            object_type_str = object_type(parse_type)
            if address_text[0] == ':':
                address_text = address_text[1:]
            #print(address_text)
            if ';' in address_text:
                streets = address_text.split(';')
            else:
                streets = [address_text]
            for street in streets:
                #print (street)
                if street != '':
                    if ',' in street or 'нечетная' in street or ' четная' in street:
                        #print (street)
                        if ',' in street:
                            street_parts = street.split(',')
                        elif ',' not in street or 'нечетная' in street or ' четная' in street:
                            street_parts = [street]
                    #else:
                    #    street_parts = [street]
                        if '–' in street_parts[0]:
                            #print(street_parts[0])
                            street_name = street_parts[0].split('–')[0]
                            street_parts[0] = street_parts[0].split('–')[1]
                        else:
                            #print(street_parts[0])
                            street_name = street_parts[0]
                            street_parts = street_parts[1:]
                        for street_part in street_parts:
                            #print(street_part)
                            odd_even_flag = 'none'
                            if 'нечетная' in street_part or ' четная' in street_part:
                                #print(street_name+street_part)
                                if 'нечетная' in street_part and ' четная' not in street_part:
                                    odd_even_flag = 'нечетная'
                                if ' четная' in street_part and 'нечетная' not in street_part:
                                    odd_even_flag = ' четная'
                                parts = street_part.partition(odd_even_flag)[2]
                                div_postfix = ''
                                begin = 0
                                end = 0
                                for number_part in parts.split(' '):
                                    if '/' in number_part:
                                        div_parts = number_part.split('/')
                                        number_part = div_parts[0]
                                        div_postfix = div_parts[1]
                                    number_part = re.sub('[^A-Za-z0-9]+', '', number_part)
                                    if number_part.isdigit() and begin == 0:
                                        begin = int(number_part.strip())
                                    if number_part.isdigit() and begin != 0:
                                        end = int(number_part.strip())
                                if begin != 0 and end != 0 and begin != end:
                                        for i in range(begin, end+1, 2):
                                            paragraphs.append(object_type_str+street_name+' '+str(i))
                                if div_postfix != '':
                                            paragraphs.append(object_type_str+street_name+' '+str(end)+'/'+div_postfix)
                                if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                        if odd_even_flag == ' четная':
                                            paragraphs.append(object_type_str+street_name+' '+str(begin)+' even to end')
                                        else:
                                            paragraphs.append(object_type_str+street_name+' '+str(begin)+' odd to end')
                                if street_part.endswith('нечетная'):
                                        #street_name = street.split('–')[0]
                                        paragraphs.append(object_type_str+street_name+' odd from start to end')
                                elif street_part.endswith(' четная'):
                                        #street_name = street.split('–')[0]
                                        paragraphs.append(object_type_str+street_name+' even from start to end')
                            else:
                                paragraphs.append(object_type_str+street_name+' '+street_part)
                    else:
                        paragraphs.append(object_type_str+street) 
with io.open('parse_result_2.csv',"w", encoding="utf-8") as textFile:
    for paragraph in paragraphs: 
        textFile.write((paragraph)+'\n')
qbfile.close()
