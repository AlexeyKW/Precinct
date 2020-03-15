#!/usr/bin/env python
# coding: utf-8

# In[2]:


import herepy
import io
import shutil
import os
import sys
import re


# https://www.shanelynn.ie/batch-geocoding-in-python-with-google-geocoding-api/

# In[3]:


geocoderApi = herepy.GeocoderApi('api-key')


# In[7]:


get_ipython().run_line_magic('time', '')

parse_file = open('parse_result_2_berdsk.csv', "r", encoding='utf-8')
paragraphs = []
ptext = ''
precinct_place_flag = 0
curr_precinct_place = ''
curr_precinct_place_lat = ''
curr_precinct_place_lon = ''
curr_precinct_number = ''
paragraphs.append('precinct_number,precinct_place,p_lat,p_lon,address,lat,lon')

i = 0
for line in parse_file:
    i=i+1
    parts = line.split(',')
    if parts[0] == 'Precinct_number':
        curr_precinct_number = parts[1].rstrip()
    if parts[0] == 'Precinct_place':
        curr_precinct_place = parts[1].rstrip()
        #print(parts[1])
        response = geocoderApi.free_form(parts[1].rstrip())
        res_dict = response.as_dict()
        try:
            curr_precinct_place_lat = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Latitude']
        except IndexError:
            curr_precinct_place_lat = 0
        try:
            curr_precinct_place_lon = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Longitude']
        except IndexError:
            curr_precinct_place_lon = 0
    if 'Address_' in parts[0]:
        #if 'even to end' in parts[1] or 'odd to end' in parts[1]:
        #    #print (parts[1])
        #    start_number = 0
        #    street_string = ''
        #    if 'even to end' in parts[1]:
        #        parts_even = parts[1].split(' ')
        #        start_number = int(parts_even[parts_even.index("even")-1])
        #        print(start_number)
        #        for x in range (0, parts_even.index("even")-1):
        #            street_string = street_string+' '+parts_even[x]
        #        print(street_string)
        #    if 'odd to end' in parts[1]:
        #        parts_odd = parts[1].split(' ')
        #        start_number = int(parts_odd[parts_even.index("odd")-1])
        #        print(start_number)                
        #        for x in range (0, parts_odd.index("odd")-1):
        #            street_string = street_string+' '+parts_odd[x]
        #        print(street_string)
        #    if start_number > 0 and street_string != '':
        #        num_zeros = 0
        #        while num_zeros<3:
        #            print('Новосибирск '+street_string+' '+str(start_number))
        #            response = geocoderApi.free_form('Новосибирск '+street_string+' '+str(start_number))
        #            res_dict = response.as_dict()
        #            lat = 0
        #            lon = 0
        #            try:
        #                lat = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Latitude']
        #            except IndexError:
        #                lat = 0
        #                num_zeros=num_zeros+1
        #            try:
        #                lon = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Longitude']
        #                paragraphs.append(str(curr_precinct_number)+','+curr_precinct_place+','+str(curr_precinct_place_lat)+','+str(curr_precinct_place_lon)+','+street_string+' '+str(start_number)+','+str(lat)+','+str(lon))
        #                #print(street_string+' '+str(start_number))
        #            except IndexError:
        #                lon = 0
        #            start_number = start_number + 2

        response = geocoderApi.free_form('Бердск '+parts[1].rstrip())
        #print(parts[1].rstrip())
        res_dict = response.as_dict()
        try:
            lat = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Latitude']
        except IndexError:
            lat = 0
        try:
            lon = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Longitude']
        except IndexError:
            lon = 0
        paragraphs.append(str(curr_precinct_number)+','+curr_precinct_place+','+str(curr_precinct_place_lat)+','+str(curr_precinct_place_lon)+','+parts[1].rstrip()+','+str(lat)+','+str(lon))
    #if parts[0] == 'Address_side':
    #    response = geocoderApi.free_form('Новосибирск переулок'+parts[1].rstrip())
        #print(parts[1].rstrip())
    #    res_dict = response.as_dict()
    #    try:
    #        lat = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Latitude']
    #    except IndexError:
    #        lat = 0
    #    try:
    #        lon = res_dict['Response']['View'][0]['Result'][0]['Location']['NavigationPosition'][0]['Longitude']
    #    except IndexError:
    #        lon = 0
    #    paragraphs.append(str(curr_precinct_number)+','+curr_precinct_place+','+str(curr_precinct_place_lat)+','+str(curr_precinct_place_lon)+','+parts[1].rstrip()+','+str(lat)+','+str(lon))
    #if i == 20:
    #    break
    if i%100 == 0:
        print(i)

with io.open('parse_result_berdsk_lat_lon.csv',"w", encoding="utf-8") as textFile:
    for paragraph in paragraphs: 
        textFile.write((paragraph)+'\n')

parse_file.close()


# In[31]:


geocoderApi.free_form('Новосибирск улица Доватора 400').as_dict()


# In[ ]:




