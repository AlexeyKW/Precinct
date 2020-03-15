#!/usr/bin/env python
# coding: utf-8

# In[1]:


from docx import Document
import io
import shutil
import os
import sys
import re


# In[2]:


path = "docx/"


# In[3]:


curr_file = ""


# In[4]:


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


# In[5]:


convertDocxToText(path)


# In[5]:


def object_type(p_type):
    object_string = 'none'
    if 'улицы' == p_type:
        object_string = 'Address_street, улица '
    elif 'Улицы' == p_type:
        object_string = 'Address_street, улица '
    elif 'Улица' == p_type:
        object_string = 'Address_street, улица '
    elif 'улица' == p_type:
        object_string = 'Address_street, улица '
    elif '– Проспект' == p_type:
        object_string = 'Address_avenue, проспект ' 
    elif 'переулок' == p_type:
        object_string = 'Address_side, переулок '
    elif 'переулки' == p_type:
        object_string = 'Address_side, переулок '
    elif 'территория' == p_type:
        object_string = 'Address_place, '
    elif 'территории' == p_type:
        object_string = 'Address_place, '
    elif '– Площадь' == p_type:
        object_string = 'Address_square, площадь '
    elif 'проезды' == p_type:
        object_string = 'Address_thoroughfare, проезд '
    elif 'тупик ' == p_type:
        object_string = 'Address_deadend, тупик '
    elif 'тупики' == p_type:
        object_string = 'Address_deadend, тупик '
    elif 'бульвар ' == p_type:
        object_string = 'Address_boulevard, бульвар '
    elif 'Гусинобродское шоссе' == p_type:
        object_string = 'Address_street, улица Гусинобродское шоссе '
    elif 'Вокзальная магистраль' == p_type:
        object_string = 'Address_street, улица Вокзальная магистраль '
    elif 'проспект Димитрова' == p_type:
        object_string = 'Address_street, проспект Димитрова '
    else:
        object_string = 'none'
    return object_string


# In[6]:


curr_file = "iskitim.txt"


# In[19]:


qbfile = open(curr_file, "r", encoding='utf-8')
paragraphs = []
ptext = ''
precinct_place_flag = 0
precinct_place_text = ''
curr_precinct_number = ''

for aline in qbfile:
    #print(aline)
    if precinct_place_flag == 1:
        if aline == "\n":
            precinct_place_flag = 0
            #print(aline)
            precinct_place = []
            if '–' in precinct_place_text:
                precinct_place = precinct_place_text.split('–')
            if ':' in precinct_place_text:
                precinct_place = precinct_place_text.split(':')
                #print(precinct_place[0])
            if '-' in precinct_place_text:
                precinct_place = precinct_place_text.split('-')
            #print(precinct_place_text)
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
        if '-' in aline:
            precinct_nums_str = aline.partition('Номера')[2].split('-')[1].strip()
        if '–' in aline:
            precinct_nums_str = aline.partition('Номера')[2].split('–')[1].strip()
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
    elif 'улица' in aline:
        parse_type = 'улица'
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
    elif 'тупик ' in aline:
        parse_type = 'тупик'
    elif 'тупики' in aline:
        parse_type = 'тупики'
    elif 'проезды' in aline:
        parse_type = 'проезды'
    elif 'бульвар ' in aline:
        parse_type = 'бульвар '
    elif 'Гусинобродское шоссе' in aline:
        parse_type = 'Гусинобродское шоссе'
    elif 'Вокзальная магистраль' in aline:
        parse_type = 'Вокзальная магистраль'
    elif 'проспект Димитрова' in aline:
        parse_type = 'проспект Димитрова'
    else:
        parse_type = 'none'
    
    if parse_type != 'none':
        address_text = aline.rpartition(parse_type)[2].rstrip()
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
                    if ',' in street or 'нечетная' in street or ' четная' in street or (' с ' in street and ' по ' in street):
                        #print (street)
                        street_parts = []
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
                            #if len(street_parts[0])>5 and street_name_parts[-1].strip().isdigit():
                            #if len(street_parts[0])>5:
                            #    #print(street_parts[0])
                            #    street_name_parts = street_parts[0].strip().split(' ')
                            #    street_name = " ".join(street_name_parts[:-1])
                            #    #print(street_name)
                            #    #print(street_name_parts[-1])
                            #    street_parts = [street_name_parts[-1]]+street_parts[1:]
                            #    #print(street_parts)
                            #else:
                            street_name = street_parts[0]
                            street_parts = street_parts[1:]
                        for street_part in street_parts:
                            #print(street_part)
                            odd_even_flag = 'none'
                            if 'нечетная' in street_part or ' четная' in street_part or (' с ' in street_part and ' по ' in street_part):
                                #print(street_name+street_part)
                                if 'нечетная' in street_part and ' четная' not in street_part:
                                    odd_even_flag = 'нечетная'
                                if ' четная' in street_part and 'нечетная' not in street_part:
                                    odd_even_flag = ' четная'
                                if odd_even_flag != 'none':
                                    parts = street_part.partition(odd_even_flag)[2]
                                else:
                                    parts = street_part
                                begin_div_postfix = ''
                                end_div_postfix = ''
                                begin = 0
                                begin_postfix = ''
                                end = 0
                                end_postfix = ''
                                for number_part in parts.split(' '):
                                    if '/' in number_part:
                                        div_parts = number_part.split('/')
                                        number_part = div_parts[0]
                                        if begin == 0:
                                            begin_div_postfix = div_parts[1]
                                        if begin != 0:
                                            end_div_postfix = div_parts[1]
                                    number_part_postfix_check = number_part
                                    number_part = re.sub('[^A-Za-z0-9]+', '', number_part)
                                    if number_part.isdigit() and begin == 0:
                                        begin = int(number_part.strip())
                                        if re.search(r'[а-яА-Я]', number_part_postfix_check[-1]):
                                            begin_postfix = number_part_postfix_check[-1]
                                    if number_part.isdigit() and begin != 0:
                                        end = int(number_part.strip())
                                        if re.search(r'[а-яА-Я]', number_part_postfix_check[-1]):
                                            end_postfix = number_part_postfix_check[-1]
                                if begin_postfix != '':
                                    paragraphs.append(object_type_str+street_name+' '+str(begin)+begin_postfix)
                                if begin_div_postfix != '':
                                    paragraphs.append(object_type_str+street_name+' '+str(begin)+'/'+begin_div_postfix)
                                if begin != 0 and end != 0 and begin != end:
                                        if odd_even_flag != 'none':
                                            for i in range(begin, end+1, 2):
                                                paragraphs.append(object_type_str+street_name+' '+str(i))
                                        else:
                                            for i in range(begin, end+1, 1):
                                                paragraphs.append(object_type_str+street_name+' '+str(i))
                                if end_div_postfix != '':
                                    paragraphs.append(object_type_str+street_name+' '+str(end)+'/'+end_div_postfix)
                                if end_postfix != '':
                                    paragraphs.append(object_type_str+street_name+' '+str(end)+end_postfix)
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
with io.open('parse_result_2_iskitim.csv',"w", encoding="utf-8") as textFile:
    for paragraph in paragraphs: 
        textFile.write((paragraph)+'\n')
qbfile.close()




