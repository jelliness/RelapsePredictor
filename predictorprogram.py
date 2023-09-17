import openpyxl

wb = openpyxl.load_workbook(r'SmokingDataSet.xlsx')
sheet = wb["Summary"]

P_age={}
P_gender={}
P_cvStatus={}
P_cessation={}
P_empStatus={}
P_type={}
P_ageStart={}
P_influence={}
P_urge={}
P_noSticks={}
P_mainAccess={}


for i in range(5):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_age[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_age)             
            
for i in range(7,9):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_gender[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_gender)          

for i in range(11,13):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_cvStatus[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_cvStatus)

for i in range(15,17):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_cessation[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_cessation)

for i in range(19,23):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_empStatus[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_empStatus)

for i in range(25,27):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_type[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_type)

for i in range(29,32):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_ageStart[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_ageStart)

for i in range(34,37):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_influence[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_influence)

for i in range(39,44):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_urge[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_urge)

for i in range(46,54):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_noSticks[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_noSticks)

for i in range(56,61):
    if str(sheet['B'+str(i+1)].value)=='None':
        continue
    P_mainAccess[str(sheet['A'+str(i+1)].value)]={
    'Total':sheet['B'+str(i+1)].value,
    'P(Yes)':sheet['C'+str(i+1)].value,
    'P(No)':sheet['D'+str(i+1)].value,
        }
    #print(P_mainAccess)


N_age='17 - 30'
N_gender='M'
N_cvStatus='Single'
N_cessation='Yes'
N_empStatus='Officeing'
N_type='SocialSmoker'
N_ageStart='15 - 19'
N_influence='PeerPressure'
N_urge='Happy'
N_noSticks=' 16 - 20'
N_mainAccess='Bars'

numerator1 =(21/69)*float(P_age[N_age]['P(Yes)'])*float(P_gender[N_gender]['P(Yes)'])*float(P_cvStatus[N_cvStatus]['P(Yes)'])*float(P_cessation[N_cessation]['P(Yes)'])*float(P_empStatus[N_empStatus]['P(Yes)'])*float(P_type[N_type]['P(Yes)'])*float(P_ageStart[N_ageStart]['P(Yes)'])*float(P_influence[N_influence]['P(Yes)'])*float(P_urge[N_urge]['P(Yes)'])*float(P_noSticks[N_noSticks]['P(Yes)'])*float(P_mainAccess[N_mainAccess]['P(Yes)'])
denominator=float(P_age[N_age]['Total'])*float(P_gender[N_gender]['Total'])*float(P_cvStatus[N_cvStatus]['Total'])*float(P_cessation[N_cessation]['Total'])*float(P_empStatus[N_empStatus]['Total'])*float(P_type[N_type]['Total'])*float(P_ageStart[N_ageStart]['Total'])*float(P_influence[N_influence]['Total'])*float(P_urge[N_urge]['Total'])*float(P_noSticks[N_noSticks]['Total'])*float(P_mainAccess[N_mainAccess]['Total'])
relapseYes=numerator1/denominator
print(numerator1)

numerator2=(48/69)*float(P_age[N_age]['P(No)'])*float(P_gender[N_gender]['P(No)'])*float(P_cvStatus[N_cvStatus]['P(No)'])*float(P_cessation[N_cessation]['P(No)'])*float(P_empStatus[N_empStatus]['P(No)'])*float(P_type[N_type]['P(No)'])*float(P_ageStart[N_ageStart]['P(No)'])*float(P_influence[N_influence]['P(No)'])*float(P_urge[N_urge]['P(No)'])*float(P_noSticks[N_noSticks]['P(No)'])*float(P_mainAccess[N_mainAccess]['P(No)'])
relapseNo=numerator2/denominator
print(numerator2)

YES=relapseYes/relapseYes+relapseNo
NO=relapseNo/relapseYes+relapseNo

print(YES)
print(NO)
