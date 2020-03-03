#!/usr/bin/env python
# coding: utf-8

# In[1]:


import herepy
import io
import shutil
import os
import sys
import re


# In[2]:


geocoderApi = herepy.GeocoderApi('api-key')


# In[1]:


get_ipython().run_line_magic('time', '')

parse_file = open('parse_result.csv', "r", encoding='utf-8')
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
    if parts[0] == 'Address_street':
        response = geocoderApi.free_form('Новосибирск '+parts[1].rstrip())
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
    if parts[0] == 'Address_side':
        response = geocoderApi.free_form('Новосибирск переулок'+parts[1].rstrip())
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
    if parts[0] == 'Address_avenue':
        response = geocoderApi.free_form('Новосибирск Проспект '+parts[1].rstrip())
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
    #if i == 200:
    #    break
    if i%1000 == 0:
        print(i)

with io.open('parse_result_lat_lon.csv',"w", encoding="utf-8") as textFile:
    for paragraph in paragraphs: 
        textFile.write((paragraph)+'\n')

parse_file.close()


# In[ ]:




