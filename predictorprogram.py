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
N_age = ""
user_input = input("Put your current age: ")
if int(user_input) < 17:
    N_age = "CHILDREN"
elif 16 < int(user_input) < 31:
    N_age = "YOUNG ADULTS"
elif 30 < int(user_input) < 46:
    N_age = "MIDDLE ADULTS"
elif int(user_input) > 45:
    N_age = "OLD ADULTS"

N_gender = "M"
user_input = input("Put your gender [M or F]: ")
if user_input == "M":
    N_gender = "M"
else:
    N_gender = "F"

N_cvStatus = ""
print("[1] - Single\n[2] - Married\n")
user_input = input("Put your civil status: ")
if user_input == "1":
    N_cvStatus = "Single"
else:
    N_cvStatus = "Married"

N_cessation = ""
user_input = input("Do you have info cessation [Y/N]: ")
if user_input == "Y":
    N_cessation = "Yes"
else:
    N_cessation = "No"

N_empStatus = ""
print("[1] - Employed\n[2] - NotOfficeing\n[3] - Officeing\n[4] - Retired")
user_input = input("What is your employee status?")
if user_input == "1":
    N_empStatus = "Employed"
elif user_input == "2":
    N_empStatus = "NotOfficeing"
elif user_input == "3":
    N_empStatus = "Officeing"
else:
    N_empStatus = "Retired"

N_type = ""
print("[1] - Regular Smoker\n[2] - Social Smoker")
user_input = input("What type of smoker are you?")
if user_input == "1":
    N_type = "RegularSmoker"
else:
    N_type = "SocialSmoker"

N_ageStart=""
user_input = input("How old were you when you started smoking? ")
if int(user_input) < 15:
    N_ageStart = "10 - 14"
elif 14 < int(user_input) < 20:
    N_ageStart = "15 - 19"
elif int(user_input) > 20:
    N_ageStart = "Above 20"

N_influence=''
print('\n[1] - Curiosity\n[2] - Family Influence\n[3] - Peer Pressure')
user_input = input("Put your smoke influence: ")
if user_input == "1":
    N_influence = "Curiosity"
elif user_input == "2":
    N_influence = "FamilyInfluence"
else:
    N_influence = "PeerPressure"

N_urge = ""
print('\n[1] - Stressed\n[2] - Bored\n[3] - Sad\n[4] - Angry\n[5] - Happy')
user_input = input("Put your urge: ")
if user_input == "1":
    N_urge = "Stressed"
elif user_input == "2":
    N_urge = "Bored"
elif user_input == "3":
    N_urge = "Sad"
elif user_input == "4":
    N_urge = "Angry"
else:
    N_urge = "Happy"

N_noSticks=''
user_input = input("Enter the number of sticks per day: ")
if int(user_input) < 6:
    N_noSticks = "1 - 5"
elif 5 < int(user_input) < 11:
    N_noSticks = "6 - 10"
elif 10 < int(user_input) < 16:
    N_noSticks = "11 - 15"
elif 15 < int(user_input) < 21:
    N_noSticks = "16 - 20"
elif 20 < int(user_input) < 26:
    N_noSticks = "21 - 25"
elif 25 < int(user_input) < 31:
    N_noSticks = "26 - 30"
elif 30 < int(user_input) < 36:
    N_noSticks = "31 - 35"
elif 35 < int(user_input) < 41:
    N_noSticks = "36 - 40"

N_mainAccess=''
print('[1] - Home\n[2] - Office\n[3] - Public Place\n[4] - Others\n[5] - Bars')
user_input = input("Enter main access: ")
if user_input == "1":
    N_mainAccess = "Home"
elif user_input == "2":
    N_mainAccess = "Office"
elif user_input == "3":
    N_mainAccess = "PublicPlace"
elif user_input == "4":
    N_mainAccess = "Others"
else:
    N_mainAccess = "Bars"

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
