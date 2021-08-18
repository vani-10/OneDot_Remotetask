import json
import xlsxwriter
import xlrd, xlwt
from xlutils.copy import copy as xl_copy

## Used for Normalizing color values
ColorMapping = {"gold":"gold","grn":"green", "schwarz":"black", "rot":"red", "silber":"silver", "anthrazit":"anthracite","weiss":"white","blau":"blue","grau":"grey","gelb":"yellow","violett":"voilet","grn":"green","orange":"orange","beige":"beige","bordeaux":"bordeaux","braun":"brown"}

## Used for Normalizing DriverType values
DriverTypeMapping= {"Allrad":"1","Hinterradantrieb":"2","Vorderradantrieb":"3","null":"0"}

## Target schema and its default value
TargetSchemaMap = {"carType":"NewSupplier","color":"Unknown","condition":"NewSupplier","currency":"NewSupplier","drive":"NewSupplier","city":"Unknown","country":"NewSupplier","make":"NewSupplier","manufacture_year":"NewSupplier","mileage":"NewSupplier","mileage_unit":"NewSupplier","model":"NewSupplier","model_variant":"NewSupplier","price_on_request":"NewSupplier","type":"NewSupplier","zip":"NewSupplier","manufacture_month":"NewSupplier","fuel_consumption_unit":"NewSupplier"}

## Utilify functions
def stripNonAlphaNum(text):
    import re
    list = re.compile(r'\W+').split(text)
    stripped_string = re.sub(r'\W+', '', list[0])
    alphanumeric = ""

    for character in stripped_string:
        if ord(character) >= 65 and ord(character) <= 90:
             alphanumeric += character
            ## checking for lower case
        elif ord(character) >= 97 and ord(character) <= 122:
             alphanumeric += character
    return alphanumeric.lower()

def mynormalize(dataType, text):
    if dataType=="color":
        germanColor = stripNonAlphaNum(text)
        return ColorMapping[germanColor]
    elif dataType=="drive":
        return  DriverTypeMapping[text]
    elif dataType=="milage":
        return  text.split(" ")[0]

## Final output in MasterExcel.xlsx
workbook = xlsxwriter.Workbook('MasterExcel.xlsx')

### Add 1st Tab/Sheet now
worksheet = workbook.add_worksheet()
row = 0
col = 0
row+=1
file1 = open('supplier_car.json', 'r')
Lines = file1.readlines()
headingAdded = False
for line in Lines:
    r = line.strip()
    rowData = json.loads(r)
    if headingAdded == False:
        for keyInSuppliersData in rowData:
            worksheet.write(row, col,keyInSuppliersData)
            col+=1
            headingAdded = True
        row+=1
    col=0
    for keyInSuppliersData in rowData:
        #print("key is {} and its valye is {}".format(keyInSuppliersData, rowData[keyInSuppliersData]))
        worksheet.write(row, col,rowData[keyInSuppliersData])
        col+=1
    row+=1

### Add 2nd Tab/Sheet now

worksheet2 = workbook.add_worksheet()
row = 0
col = 0
row+=1
headingAdded = False
for line in Lines:
    r = line.strip()
    rowData = json.loads(r)
    if headingAdded == False:
        for keyInSuppliersData in rowData:
            worksheet2.write(row, col,keyInSuppliersData)
            col+=1
            headingAdded = True
        row+=1
    col=0
    for keyInSuppliersData in rowData:
        if keyInSuppliersData == "Attribute Names" and rowData[keyInSuppliersData] == "BodyColorText":
            #print("in normalize")
            worksheet2.write(row, col,rowData[keyInSuppliersData])
            col+=1
            suppliedBodyColr =  mynormalize("color", rowData["Attribute Values"])
            #print("val is {}".format(suppliedBodyColr))
            worksheet2.write(row, col,suppliedBodyColr)
            col+=1
        elif keyInSuppliersData == "Attribute Values" and rowData["Attribute Names"] == "BodyColorText":
            #print("skipped")
            continue
        elif keyInSuppliersData == "Attribute Names" and rowData[keyInSuppliersData] == "DriveTypeText":
            worksheet2.write(row, col,rowData[keyInSuppliersData])
            col+=1
            normalizedDriverType = mynormalize("drive", rowData["Attribute Values"])
            worksheet2.write(row, col, normalizedDriverType)
            col+=1
        elif keyInSuppliersData == "Attribute Values" and rowData["Attribute Names"] == "DriveTypeText":
            #print("skipped")
            continue
        elif keyInSuppliersData == "Attribute Names" and rowData[keyInSuppliersData] == "ConsumptionTotalText":
            worksheet2.write(row, col,rowData[keyInSuppliersData])
            col+=1
            normalizedMilage = mynormalize("milage", rowData["Attribute Values"])
            worksheet2.write(row, col, normalizedMilage)
            col+=1
        elif keyInSuppliersData == "Attribute Values" and rowData["Attribute Names"] == "ConsumptionTotalText":
            #print("skipped")
            continue
        else:
            worksheet2.write(row, col,rowData[keyInSuppliersData])
            col+=1
    row+=1

### Add 3rd Tab/Sheet now

worksheet3 = workbook.add_worksheet()
row = 0
col = 0

# Add the Target schema header into excel
for keyInTargetschema in TargetSchemaMap:
    worksheet3.write(row, col, keyInTargetschema)
    col+=1

row+=1
col=0
for line in Lines:
    r = line.strip()
    rowData = json.loads(r)
    
    # populate local object with only required values (like color, city, make, model and model_variant) from rowData (i.e. supplier data)
    localRetrievedData = {}
    for keyInSuppliersData in rowData:
        #print("key is {} and its value is {}".format(keyInSuppliersData, rowData[keyInSuppliersData]))        
        if keyInSuppliersData == "Attribute Names" and rowData[keyInSuppliersData] == "BodyColorText":
            localRetrievedData["color"]=mynormalize("color", rowData["Attribute Values"])
        if keyInSuppliersData == "Attribute Names" and rowData[keyInSuppliersData] == "ConsumptionTotalText":
            localRetrievedData["mileage"]=mynormalize("milage", rowData["Attribute Values"])
        elif keyInSuppliersData == "Attribute Names" and rowData[keyInSuppliersData] == "City":
            localRetrievedData["city"]=rowData["Attribute Values"]
        elif keyInSuppliersData == "MakeText":
            localRetrievedData["make"]=rowData[keyInSuppliersData]
        elif keyInSuppliersData == "ModelText":
            localRetrievedData["model"]=rowData[keyInSuppliersData]
        elif keyInSuppliersData == "ModelTypeText":
            localRetrievedData["model_variant"]=rowData[keyInSuppliersData]

    # Now fill the worsheet with values localRetrievedData, if exists. Else add default values defined in TargetSchemaMap.
    col=0
    for keyInTargetschema in TargetSchemaMap:
        if keyInTargetschema in localRetrievedData:
            worksheet3.write(row, col,localRetrievedData[keyInTargetschema])
        else:
            worksheet3.write(row, col,TargetSchemaMap[keyInTargetschema])     
        col+=1

    row+=1

## close the workbook. This is the end.
workbook.close()