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


# In[7]:


convertDocxToText(path)


# In[9]:


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
            paragraphs.append('Precinct_address,'+precinct_place[1].replace(')', '').partition('»,')[2])
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
                #print(street)
                if ',' in street:
                    street_parts = street.split(',')
                    street_name = street_parts[0]
                    if '–' in street_parts[0]:
                        street_name = street_parts[0].split('–')[0]
                        #print(street_parts[0])
                        if 'нечетная' in street_parts[0].split('–')[1] or 'четная' in street_parts[0].split('–')[1]:
                                    if 'нечетная' in street_parts[0].split('–')[1] and ' четная' not in street_parts[0].split('–')[1]:
                                        begin = 0
                                        end = 0
                                        odd_parts = street_parts[0].split('–')[1].partition('нечетная')
                                        div_postfix = ''
                                        #street_name = odd_parts[0]
                                        for number_part in odd_parts[2].split(' '):
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
                                                paragraphs.append('Address_street,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_street,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_street,'+street_name+' '+str(begin)+' odd to end')
                                        #if begin == 0 and end == 0 and street_parts[0].endswith('нечетная'):
                                        #    print(street_parts[0])
                                        #    paragraphs.append('Address_street,'+street_name+' '+' odd from start to end')
                                        
                                    if ' четная' in street_parts[0].split('–')[1] and 'нечетная' not in street_parts[0].split('–')[1]:
                                        begin = 0
                                        end = 0
                                        even_parts = street_parts[0].split('–')[1].partition(' четная')
                                        #print(even_parts[2])
                                        #street_name = even_parts[0]
                                        div_postfix = ''
                                        for number_part in even_parts[2].split(' '):
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
                                                paragraphs.append('Address_street,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_street,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_street,'+street_name+' '+str(begin)+' even to end')
                                        #if begin == 0 and end == 0 and street_parts[0].endswith('четная'):
                                        #    print(street_parts[0])
                                        #    paragraphs.append('Address_street,'+street_name+' '+' even from start to end')
                        
                    for street_part in street_parts:
                            if street_part != street_parts[0]:
                                if 'нечетная' in street_part or 'четная' in street_part:
                                    #print()
                                    if 'нечетная' in street_part and ' четная' not in street_part:
                                        begin = 0
                                        end = 0
                                        odd_parts = street_part.partition('нечетная')
                                        #street_name = odd_parts[0]
                                        div_postfix = ''
                                        for number_part in odd_parts[2].split(' '):
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
                                                paragraphs.append('Address_street,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_street,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_street,'+street_name+' '+str(begin)+' odd to end')
                                            #print(street_parts[0])

                                    if ' четная' in street_part and 'нечетная' not in street_part:
                                        begin = 0
                                        end = 0
                                        even_parts = street_part.partition(' четная')
                                        #print(even_parts[2])
                                        #street_name = even_parts[0]
                                        div_postfix = ''
                                        for number_part in even_parts[2].split(' '):
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
                                                paragraphs.append('Address_street,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_street,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_street,'+street_name+' '+str(begin)+' even to end')
                                else:
                                    paragraphs.append('Address_street,'+street_parts[0]+street_part)
                else:
                    if street.endswith('нечетная'):
                        street_name = street.split('–')[0]
                        paragraphs.append('Address_street,'+street_name+' odd from start to end')
                    elif street.endswith(' четная'):
                        street_name = street.split('–')[0]
                        paragraphs.append('Address_street,'+street_name+' even from start to end')
                    else:
                        if 'нечетная' in street or 'четная' in street:
                            street_name = street.split('–')[0]
                                    #print()
                            if 'нечетная' in street and ' четная' not in street:
                                        begin = 0
                                        end = 0
                                        odd_parts = street.partition('нечетная')
                                        #street_name = odd_parts[0]
                                        div_postfix = ''
                                        for number_part in odd_parts[2].split(' '):
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
                                                paragraphs.append('Address_street,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_street,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_street,'+street_name+' '+str(begin)+' odd to end')
                                            #print(street_parts[0])

                            if ' четная' in street and 'нечетная' not in street:
                                        begin = 0
                                        end = 0
                                        even_parts = street.partition(' четная')
                                        #print(even_parts[2])
                                        #street_name = even_parts[0]
                                        div_postfix = ''
                                        for number_part in even_parts[2].split(' '):
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
                                                paragraphs.append('Address_street,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_street,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_street,'+street_name+' '+str(begin)+' even to end')
                        else:
                            paragraphs.append('Address_street,'+street)
        elif 'Улица' in aline:
            address_text = aline.partition('Улица')[2].rstrip()
            if ',' in address_text:
                #print(aline)
                street_parts = address_text.split(',')
                for street_part in street_parts:
                    if street_part != street_parts[0]:
                        #print(street_part)
                        paragraphs.append('Address_street,'+street_parts[0]+street_part)
        elif 'Границы участка – Проспект' in aline:
            address_text = aline.partition('Проспект')[2].rstrip()
            if ',' in address_text:
                #print(aline)
                street_parts = address_text.split(',')
                for street_part in street_parts:
                    if street_part != street_parts[0]:
                        if '–' in street_part:
                            street_name = street_part.split('–')[0]
                            if 'нечетная' in street_part.split('–')[1] or 'четная' in street_part.split('–')[1]:
                                    if 'нечетная' in street_part.split('–')[1] and ' четная' not in street_part.split('–')[1]:
                                        begin = 0
                                        end = 0
                                        odd_parts = street_part.split('–')[1].partition('нечетная')
                                        div_postfix = ''
                                        #street_name = odd_parts[0]
                                        for number_part in odd_parts[2].split(' '):
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
                                                paragraphs.append('Address_avenue,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_avenue,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_avenue,'+street_name+' '+str(begin)+' odd to end')
                                        #if begin == 0 and end == 0 and street_parts[0].endswith('нечетная'):
                                        #    print(street_parts[0])
                                        #    paragraphs.append('Address_street,'+street_name+' '+' odd from start to end')
                                        
                                    if ' четная' in street_parts[0].split('–')[1] and 'нечетная' not in street_parts[0].split('–')[1]:
                                        begin = 0
                                        end = 0
                                        even_parts = street_parts[0].split('–')[1].partition(' четная')
                                        #print(even_parts[2])
                                        #street_name = even_parts[0]
                                        div_postfix = ''
                                        for number_part in even_parts[2].split(' '):
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
                                                paragraphs.append('Address_avenue,'+street_name+' '+str(i))
                                        if div_postfix != '':
                                            paragraphs.append('Address_avenue,'+street_name+' '+str(end)+'/'+div_postfix)
                                        if begin != 0 and (end == 0 or begin==end):
                                            #for i in range(begin, end+1):
                                            paragraphs.append('Address_avenue,'+street_name+' '+str(begin)+' even to end')
                                        #if begin == 0 and end == 0 and street_parts[0].endswith('четная'):
                                        #    print(street_parts[0])
                                        #    paragraphs.append('Address_street,'+street_name+' '+' even from start to end')
                        #print(street_part)
                        paragraphs.append('Address_avenue,'+street_parts[0]+street_part)
        else:
            address_text = aline.split('–')[1]
            
            if ',' in address_text:
                street_parts = street.split(',')
                for street_part in street_parts:
                    if street_part != street_parts[0]:
                        paragraphs.append('Address_street,'+street_parts[0]+' '+street_part)
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




