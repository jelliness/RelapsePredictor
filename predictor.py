import openpyxl

wb = openpyxl.load_workbook(r'SmokingDataSet (3).xlsx')
sheet = wb["Attempt2"]

def getDataFromSheet(target):
    extractedData=[]
    for i in range(1, sheet.max_row + 1):
        cell_obj = sheet.cell(row = i, column = 1)
        if str(cell_obj.value)==target:
            for j in range(1, sheet.max_column + 1):
                cell_obj = sheet.cell(row = i, column = j)
                extractedData.append(cell_obj.value)
    return extractedData

def squeeze(target):
    record=[]
    for data in target:
        record.append(getDataFromSheet(data))
    return record

def getValues(record,index):
    allVal=[]
    for group in record:
        allVal.append(group[index])
    return allVal

def calculateLaplace(num,total,a,K):
    laplace=(num+a)/(total+a*K)
    return laplace

def checkZero(list):
    laplaceExist=False
    if 0 in list:
        laplaceExist=True
    else:
        laplaceExist=False
    return laplaceExist

def getFinalValues(allVal,laplaceExist,total,a,K):
    finalVal=[]
    
    if laplaceExist:
        for num in allVal:
            partialNumerator=calculateLaplace(num,total,a,K)
            finalVal.append(partialNumerator)
    else:
        for num in allVal:
            partialNumerator=(num)/(total)
            finalVal.append(partialNumerator)

    print(finalVal)
    return finalVal

def getP_target(allVal,total,overall,a,K):
    laplaceExist=checkZero(allVal)
    p_yes=0
    if laplaceExist:
        p_yes=calculateLaplace(total,overall,a,K)
    else:
        p_yes=(total)/(overall) 
    return p_yes

def getNumerator(record,index,numOfResponse,overall,a,K):
    numerator=1
    rawNum=getValues(record,index)
    laplaceExist=checkZero(rawNum)
    print(laplaceExist)
    #laplaceExist=True
    p_target=getP_target(rawNum,numOfResponse,overall,a,K)
    partialNumerator=getFinalValues(rawNum,laplaceExist,numOfResponse,a,K)
    for num in partialNumerator:
        numerator=numerator*num
    product=numerator*p_target
    return product

def getDenominator(record,overall,checkIndex,a,K,i=3):
    denominator=1
    checkList=getValues(record,checkIndex)
    laplaceExist=checkZero(checkList)
    print(laplaceExist)
    #laplaceExist=True
    rawDenom=getValues(record,i)
    partialDenominator=getFinalValues(rawDenom,laplaceExist,overall,a,K)
    for num in partialDenominator:
        denominator=denominator*num
    return denominator

def calculateResponse(record,numOfResponse,overallNumOfData,a,K,index):
    predictedValue=0
    numerator=getNumerator(record,index,numOfResponse,overallNumOfData,a,K)
    denominator=getDenominator(record,overallNumOfData,index,a,K,i=3)
    print(numerator,"/",denominator)
    predictedValue=numerator/denominator    
    return predictedValue



target=['YOUNG ADULTS','M','Single','No','Officeing','RegularSmoker','15 - 19','FamilyInfluence','Sad','6 - 10','Bars']
a=1
K=12
totalOfYesResponse=21
totalOfNoResponse=48
overallNumOfData=totalOfYesResponse+totalOfNoResponse

record=squeeze(target)

RelapsePositive=calculateResponse(record,totalOfYesResponse,overallNumOfData,a,K,index=1)
RelapseNegative=calculateResponse(record,totalOfNoResponse,overallNumOfData,a,K,index=2)

overallTotal=RelapsePositive+RelapseNegative
negativeRelapseRate=RelapseNegative/overallTotal
positiveRelapseRate=RelapsePositive/overallTotal
totalPercentage=positiveRelapseRate+negativeRelapseRate
print()

print("P(Yes | Record): ",RelapsePositive)
print("P(No | Record): ",RelapseNegative)
print("Positive to Relapse Rate %: ", positiveRelapseRate)
print("Negative to Relapse Rate %: ", negativeRelapseRate)
print("Total Percentage %: ", totalPercentage)


