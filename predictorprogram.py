import openpyxl

wb = openpyxl.load_workbook(r'SmokingDataSet1.xlsx')
sheet = wb["SummaryOfData"]

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


N_age="YOUNG ADULTS"
N_gender='M'
N_cvStatus='Single'
N_cessation='Yes'
N_empStatus='Officeing'
N_type='SocialSmoker'
N_ageStart='15 - 19'
N_influence='PeerPressure'
N_urge='Happy'
N_noSticks='16 - 20'
N_mainAccess='Bars'

num1=P_age[N_age]['P(Yes)']*P_gender[N_gender]['P(Yes)']*P_cvStatus[N_cvStatus]['P(Yes)']*P_cessation[N_cessation]['P(Yes)']*P_empStatus[N_empStatus]['P(Yes)']*P_type[N_type]['P(Yes)']*P_ageStart[N_ageStart]['P(Yes)']*P_influence[N_influence]['P(Yes)']*P_urge[N_urge]['P(Yes)']*P_noSticks[N_noSticks]['P(Yes)']*P_mainAccess[N_mainAccess]['P(Yes)']*21/69
num2=P_age[N_age]['P(No)']*P_gender[N_gender]['P(No)']*P_cvStatus[N_cvStatus]['P(No)']*P_cessation[N_cessation]['P(No)']*P_empStatus[N_empStatus]['P(No)']*P_type[N_type]['P(No)']*P_ageStart[N_ageStart]['P(No)']*P_influence[N_influence]['P(No)']*P_urge[N_urge]['P(No)']*P_noSticks[N_noSticks]['P(No)']*P_mainAccess[N_mainAccess]['P(No)']*48/69
den1=P_age[N_age]['Total']*P_gender[N_gender]['Total']*P_cvStatus[N_cvStatus]['Total']*P_cessation[N_cessation]['Total']*P_empStatus[N_empStatus]['Total']*P_type[N_type]['Total']*P_ageStart[N_ageStart]['Total']*P_influence[N_influence]['Total']*P_urge[N_urge]['Total']*P_noSticks[N_noSticks]['Total']*P_mainAccess[N_mainAccess]['Total']
relapseYes=num1/den1
relapseNo=num2/den1


yesPercent=relapseYes/(relapseYes+relapseNo)
noPercent=relapseNo/(relapseYes+relapseNo)

print(yesPercent)
print(noPercent)
print(yesPercent+noPercent)
