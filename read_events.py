#--*-- coding: utf-8 --*--

import csv
import zoneinfo
import datetime
import pytz

skip_until_header=True
hmap={}

hki = pytz.timezone("Europe/Helsinki")

with open(r"\Users\EinoMakitalo\Downloads\koinly_2022.csv","r") as csvfile:
    for row in csv.reader(csvfile,dialect='excel'):        
        if skip_until_header:
            if len(row)==20 and row[0]=='Date':
                skip_until_header=False
                ind=0
                for cell in row:
                    hmap[cell]=ind
                    ind=ind+1                    
        else:
            utctime=pytz.utc.localize(datetime.datetime.strptime(row[0],'%Y-%m-%d %H:%M:%S UTC'))
            #print(utctime)
            x1=hki.normalize(utctime)
            print(utctime, x1)