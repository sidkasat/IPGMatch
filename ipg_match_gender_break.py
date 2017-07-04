from __future__ import print_function
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import json
import csv
import numpy as np
import math


SCOPE = ["https://spreadsheets.google.com/feeds"]
SECRETS_FILE = "IPGMatch-7e51d25345b0.json"
SPREADSHEET = "IPG 2017-18 Mentor Interest Form (Responses)"

MENTORS={'Sid Kasat':(['Donald Bren School of Information and Computer Sciences','Outdoors','Research opportunities','Male'],[],[],[],[]),
         'Apurva Jakhanwal':(['Donald Bren School of Information and Computer Sciences','Outdoors','School related Clubs','Female'],[],[],[],[]) ,
         'William Floyd':(['The Henry Samueli School of Engineering','Coding and Gaming','Greek Life- Fraternities & Sororities','Male'],[],[],[],[]),
         'Fei':(['Francisco J. Ayala School of Biological Sciences','Sports','UCI Sports Teams','Female'],[],[],[],[]),
         'Suraj Patil':(['School of Social Ecology','TV & Cinema','Greek Life- Fraternities & Sororities','Male'],[],[],[],[]),
         'Steven Jiang':(['School of Physical Sciences','Cooking','Student Body Government (ASUCI)','Male'],[],[],[],[]),
         'Wilson Xie':(['School of Humanities','Dance and Music',"Haven't thought about it yet",'Male'],[],[],[],[]),
         'Female Mentor':(['The Paul Merage School of Business','Community Service',"Haven't thought about it yet",'Female'],[],[],[],[])}

json_key = json.load(open(SECRETS_FILE))
# Authenticate using the signed key
credentials = ServiceAccountCredentials.from_json_keyfile_name(SECRETS_FILE, SCOPE)

gc = gspread.authorize(credentials)
#print("The following sheets are available")
#for sheet in gc.openall():
#    print("{} - {}".format(sheet.title, sheet.id))

ofile  = open('results.csv', "wt")
writer = csv.writer(ofile)
writer.writerow(("MATCHING","RESULTS"))

workbook = gc.open(SPREADSHEET)
sheet = workbook.sheet1
data = pd.DataFrame(sheet.get_all_records()).values
MAX_PEOPLE = math.ceil(len(data)/len(MENTORS))

print("-------------EXCEL DATA----------------")
data=data.tolist()
data=list(reversed(data))
#for i in data:
    #print(i)
print("-------------MENTORS DATA---------------")
for name,j in MENTORS.items(): 
    perfect=[]          #All 4 criteria match
    good=[]             #3 criterias match
    fair=[]             #2 criterias match
    poor=[]             #Only 1 criteria matches
    for i in reversed(data):
        score=0
        list1=[i[5],i[6],i[2],i[8]]                         #School,Hobbies,Interests,Gender
        Keyword=list1+j[0]
        common = list((set(list1) & set(j[0])) & set(Keyword))
        #print("COMMON: ",common)                           #Prints all common criteria between mentor and mentee
        #print(tup1,j)
        if(i[8]!=j[0][3] and i[8]!='No preference'):
            continue
        
            
        if(len(common)==3 and (len(j[1])<MAX_PEOPLE)):      #All 3 criteria same- PERFECT
            j[1].append(i[7])        
            perfect.append(i[7])
            print("----Deleting : ",i[7])
            data.remove(i)
            for key,val in MENTORS.items():
                if(name!=key and i[7] in MENTORS[key][2]):
                    MENTORS[key][2].remove(i[7])
                elif(name!=key and i[7] in MENTORS[key][3]):
                    MENTORS[key][3].remove(i[7])
                elif(name!=key and i[7] in MENTORS[key][4]):
                    MENTORS[key][4].remove(i[7])
        if(len(j[1])==MAX_PEOPLE):
           del j[2][:]
           del j[3][:]
           del j[4][:]
        elif(len(common)==3):                       #Any 2 criteria same- GOOD
            good.append(i[7])
            j[2].append(i[7])
            for key,val in MENTORS.items():
                if(name!=key and i[7] in MENTORS[key][3]):
                    MENTORS[key][3].remove(i[7])
                elif(name!=key and i[7] in MENTORS[key][4]):
                    MENTORS[key][4].remove(i[7])
        elif(len(common)==2):                       #Any 1 criteria same- FAIR
            fair.append(i[7])
            j[3].append(i[7])
            for key,val in MENTORS.items():
                if(name!=key and i[7] in MENTORS[key][4]):
                    MENTORS[key][4].remove(i[7])
        else:                                       #None same- POOR
            poor.append(i[7])
            j[4].append(i[7])
        
        #if list1 == j:
        #    print("Perfect match between ",i[6]," and ",name)
    '''
    print(name, "matches:")
    print("PERFECT : ", ','.join(str(o) for o in perfect))
    print("GOOD : ",','.join(str(o) for o in good))
    print("FAIR : ",','.join(str(o) for o in fair))
    print("POOR : ",','.join(str(o) for o in poor))
    print()
    
    writer.writerow((name, "matches:"))
    writer.writerow(("PERFECT : ", ','.join(str(o) for o in perfect)))
    writer.writerow(("GOOD : ",','.join(str(o) for o in good)))
    writer.writerow(("FAIR : ",','.join(str(o) for o in fair)))
    writer.writerow(("POOR : ",','.join(str(o) for o in poor)))
    writer.writerow(())
    '''
for key,val in MENTORS.items():
    for i in val[1]:
        for j,k in MENTORS.items():
            if (i in k[2]):
                MENTORS[j][2].remove(i)
            if (i in k[3]):
                MENTORS[j][3].remove(i)
            if (i in k[4]):
                MENTORS[j][4].remove(i)
                
for key,val in MENTORS.items():
    for i in val[2]:
        for j,k in MENTORS.items():
            if (i in k[3]):
                MENTORS[j][3].remove(i)
            if (i in k[4]):
                MENTORS[j][4].remove(i)

for key,val in MENTORS.items():
    for i in val[3]:
        for j,k in MENTORS.items():
            if (i in k[4]):
                MENTORS[j][4].remove(i)

for i,k in MENTORS.items():
    print(i, " matches :- ")
    print("PERFECT : ", ','.join(str(o) for o in k[1]))
    print("GOOD : ",','.join(str(o) for o in k[2]))
    print("FAIR : ",','.join(str(o) for o in k[3]))
    print("POOR : ",','.join(str(o) for o in k[4]))
    writer.writerow((i, "matches:"))
    writer.writerow(("PERFECT : ", ','.join(str(o) for o in k[1])))
    writer.writerow(("GOOD : ",','.join(str(o) for o in k[2])))
    writer.writerow(("FAIR : ",','.join(str(o) for o in k[3])))
    writer.writerow(("POOR : ",','.join(str(o) for o in k[4])))
    writer.writerow(())
    
ofile.close()

    
'''
#0- EMAIL
#1- Timestamp
#2- Interests at UCI
#3- Phone number
#4- Student ID
#5- School/Major
#6- Hobbies
#7- Name
#8- Gender Preference
'''
