from func import *
#param canOthers just need to 1 define
#if you define it on first param the next will be skipped
itemTest = [
       {
              "ItemName" : "Flat Call SLI 007 (Internasional call menggunakan prefix 007)",
              "EventName" : "Flat Call SLI 007",
              "FunctionName" : "exportExcelFlatCallSLI",
              "Params": {
                     "Price Pre Paid" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "false",
                            "dataType" : "int",
                            "delimiter" : "",
                            "canOthers" : "false",
                     },
                     "Price Post Paid" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "false",
                            "dataType" : "int",
                            "delimiter" : "",
                            "canOthers" : "false",
                     },
                     "Country" : {
                            "type" : "option",
                            "multiple" :"false",
                            "canOthers" : "true",
                            "dataType" : "array",
                            "value" : [
                                   'Amerika Serikat +1',
                                   'Bangladesh +880',
                                   'Brunei Darussalam +673',
                                   'China +86',
                                   'Hong Kong +852',
                                   'India +91',
                                   'Korea Selatan +82',
                                   'Malaysia +60',
                                   'Singapura +65',
                                   'Jepang +81',
                                   'Thailand +66'
                            ],
                     }, 
              }
       },
       {
              "ItemName" : "Empty Offer",
              "EventName" : "Empty Offer",
              "FunctionName" : "exportExcelEmptyOffer",
              "Params": {
                     "OfferName" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "true",
                            "dataType" : "array",
                     },
                     "OfferDesc" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     }
              }
       },
       {
              "ItemName" : "RC Offer",
              "EventName" : "RC Offer",
              "FunctionName" : "exportExcelRcOffer",
              "Params": {
                     "OfferName" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "true",
                            "dataType" : "array",
                     },
                     "OfferDesc" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Rate" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Proration" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                            "details" : "Prorate/Non Prorate",
                     },
                     "AMDD Charge Code" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
              },
       },
       {
              "ItemName" : "OC Offer",
              "EventName" : "OC Offer",
              "FunctionName" : "exportExcelOCOffer",
              "Params": {
                     "OC Type": {
                            "type" : "option",
                            "multiple" :"false",
                            "canOthers" : "false",
                            "dataType" : "array",
                            "value" : [
                                   'One Time Charge',
                                   'Installment',
                                   'Penalty',
                            ],
                     },
                     "OfferName" : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "true",
                            "dataType"    : "array",
                     },
                     "OfferDesc" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Rate" : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                     },
                     "AMDD Charge Code" : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                     },
                     "Commitment Period" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array", 
                            "conditions"  : ["Installment", "Penalty"],
                            "determiningParam" : "OC Type"
                     },
                     "AMDD Charge Code Remove Offer" : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : ["Penalty"],
                            "determiningParam" : "OC Type"
                     },
              }
       },
       {
              "ItemName" : "Discount Offer",
              "EventName" : "Discount Offer",
              "FunctionName" : "exportExcelDiscountOffer",
              "Params": {
                     "Offer Name" : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "true",
                            "dataType"    : "array",
                     },
                     "Offer Description" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Rate/Discount Value" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Rate Offer OC" : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "details"     : "Case >,<,Remove Offer",
                            "status"      : "Disabled"
                     },
                     "Rate Offer RC" : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "details"     : "Case >,<,Remove Offer",
                            "status"      : "Disabled"
                     },
              }
       },
       {
              "ItemName" : "New Service filter and new charge code",
              "EventName" : "New Service filter and new charge code",
              "FunctionName" : "exportExcelNewServiceFilter",
              "Params": {
                     "Product Name" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "true",
                            "dataType" : "array",
                     },
                     "Service Filter" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                            "details" : "ON,OFF",
                     },
                     "Charge Code" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                            "details" : "ON,OFF",
                     },
              }
       },
       {
              "ItemName" : "Minimum Usage",
              "EventName" : "Minimum Usage",
              "FunctionName" : "exportExcelMinimunUsage",
              "Params": {
                     "Offer Name" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "true",
                            "dataType" : "array",
                     },
                     "Offer Description" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Amount" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Vascode" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
              }
       },
       {
              "ItemName" : "MU BA Level",
              "EventName" : "MU BA Level",
              "FunctionName" : "exportExcelMUBALevel",
              "Params": {
                     "Offer Name" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "true",
                            "dataType" : "array",
                     },
                     "Offer Description" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Amount MU" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                     },
                     "Amount For Flexible RC" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "status"      : "Disabled"
                     },
                     "Amount For Flexible OC" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "status"      : "Disabled"
                     },
                     "Vascode" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
              }
       },
       {
              "ItemName" : "Update AMDD Charge Code",
              "EventName" : "Update AMDD Charge Code",
              "FunctionName" : "exportExcelUpdateAMDD",
              "Params": {
                     "Offer Name" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "true",
                            "dataType" : "array",
                     },
                     "Offer Description" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Offer ID" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
                     "Service Filter" : {
                            "type" : "input",
                            "value" : [
                            ],
                            "multiple" : "true",
                            "delimiter" : "|",
                            "canOthers" : "false",
                            "dataType" : "array",
                     },
              }
       },
       {
              "ItemName" : "New PP",
              "EventName" : "New PP",
              "FunctionName" : "exportExcelNewPP",
              "Params": {
                     "Type": {
                            "type" : "option",
                            "multiple" :"false",
                            "canOthers" : "false",
                            "dataType" : "array",
                            "value" : [
                                   'Prepaid',
                                   'Postpaid',
                            ],
                     },
                     "PP Name" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "false",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array", 
                     },
                     "Case (Prepaid)" : {
                            "type" : "option",
                            "multiple" :"false",
                            "canOthers" : "false",
                            "dataType" : "array",
                            "value" : [
                                   'Preload Bonus',
                                   'Tarif Voice',
                                   'Tarif SMS',
                                   'GPRS & Landing Page',
                                   'Split Recharge',
                                   'Attach offer & diameter',
                                   'Change PP',
                                   'PAYU Capped GPRSROAM'
                            ],
                            "conditions"  : ["Prepaid"],
                            "determiningParam" : "Type"
                     },
                     "Case (Postpaid)" : {
                            "type" : "option",
                            "multiple" :"false",
                            "canOthers" : "false",
                            "dataType" : "array",
                            "value" : [
                                   'Tarif Voice',
                                   'Tarif SMS',
                                   'GPRS & Landing Page',
                                   'PAYU Capped GPRSROAM',
                                   'Attach offer RC, OC, and Flexible Offer',
                                   'Change PP',
                                   'Offline Event',
                            ],
                            "conditions"  : ["Postpaid"],
                            "determiningParam" : "Type"
                     },
                     "Rounded (Time Unit)" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "true",
                            "dataType"    : "array",
                            "conditions"  : ["Tarif Voice"],
                            "determiningParam" : "Case (Postpaid)" 
                     },
                     "Rate Onnet" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Postpaid)" : ["Tarif Voice","Tarif SMS"],
                                   "Case (Prepaid)"  : ["Tarif Voice"],
                            },
                            "determiningParam" : ["Case (Postpaid)","Case (Prepaid)"] 
                     },
                     "Rate Offnet" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Postpaid)" : ["Tarif Voice","Tarif SMS"],
                                   "Case (Prepaid)"  : ["Tarif Voice"],
                            },
                            "determiningParam" : ["Case (Postpaid)"] 
                     },
                     "Rate Lokal PSTN" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Postpaid)" : ["Tarif Voice"]
                            },
                            "determiningParam" : ["Case (Postpaid)"] 
                     },
                     "PP To" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Postpaid)" : ["Change PP"],
                                   "Case (Prepaid)" : ["Change PP"]
                            },
                            "determiningParam" : ["Case (Postpaid)", "Case (Prepaid)"] 
                     },
                     "Unit" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Postpaid)" : ["GPRS & Landing Page"]
                            },
                            "determiningParam" : ["Case (Postpaid)"] 
                     },
                     "Rate" : {
                           "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "details"     : "Per Unit",
                            "conditions"  : {
                                   "Case (Postpaid)" : ["GPRS & Landing Page"]
                            },
                            "determiningParam" : ["Case (Postpaid)"] 
                     },
                     'Welcome Message'  : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Prepaid)" : ["Preload Bonus","Split Recharge"]
                            },
                            "determiningParam" : ["Case (Prepaid)"] 
                     },
                     'Preload Bonus'  : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Prepaid)" : ["Preload Bonus", "Change PP"]
                            },
                            "determiningParam" : ["Case (Prepaid)"] 
                     },
                     'Preload Balance'  : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Prepaid)" : ["Preload Bonus"]
                            },
                            "determiningParam" : ["Case (Prepaid)"] 
                     },
                     'Validity'  : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Prepaid)" : ["Preload Bonus"]
                            },
                            "determiningParam" : ["Case (Prepaid)"] 
                     },
                     'Preload Bonus (For PP TO)'  : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "conditions"  : {
                                   "Case (Prepaid)" : ["Preload Bonus", "Change PP"]
                            },
                            "determiningParam" : ["Case (Prepaid)"] 
                     },
                     'Rate Per Zone' : {
                            "type"        : "input",
                            "value"       : [
                            ],
                            "multiple"    : "true",
                            "delimiter"   : "|",
                            "canOthers"   : "false",
                            "dataType"    : "array",
                            "details"     : "A;B;C;D;E;F;G;H;Special",
                            "conditions"  : {
                                   "Case (Prepaid)" : ["Tarif SMS"]
                            },
                            "determiningParam" : ["Case (Prepaid)"] 
                     }
              }
       }
]

print("\nAvailable Item Test:\n")
for no, case in enumerate(itemTest):
    no = no + 1
    print(str(no) + " : " + str(case['ItemName']))

testInput = input("\nPlease input your item test number : ")

indexFunction = int(testInput)-1
function_name = itemTest[indexFunction]['FunctionName']
function = globals().get(function_name)
choosedTestCase = itemTest[indexFunction]
neededParams = list(itemTest[indexFunction]['Params'].keys())
params = {}
multiParam = []

if function is not None and callable(function):
       print("\nYou choose test case: " + str(choosedTestCase['ItemName']) + "\n")
       print("You need to fill the params: ")

       multiParam = []
       stopLoop = False
       arrayParam = {}
       continueCondition = False
       for key, param in choosedTestCase['Params'].items():
              if "status" in param and param['status'] == 'Disabled':
                     continue
              
              if "conditions" in param and isinstance(param['determiningParam'], str):
                     if param['determiningParam'] in params:
                            if "conditions" in param and params[param['determiningParam']][0] not in param["conditions"]:
                                   continue  # Skip the current iteration and move to the next one
                     else:
                            continue
              
              if "determiningParam" in param and isinstance(param['determiningParam'], list):
                     for determiningParam in param['determiningParam']:
                            if determiningParam in params:
                                   if params[determiningParam][0] not in param["conditions"][determiningParam]:
                                          continueCondition = True
                                          break  # Skip the current iteration and move to the next one      
                                   else:
                                          continueCondition = False
                                          break
                            else:
                                   continueCondition = True
                                   continue    
              
              if continueCondition:
                     continueCondition = False
                     continue

              details = ''
              if "details" in param:
                     details = "("+param['details']+")"
              print("\n" + str(key) + " "+details+": ")

              if param['multiple'] == 'true':
                     print("You can use multiple values with '" + param['delimiter'] + "' delimiter")

              if param['type'] == 'option':
                     optionBefore = param['value'][:]
                     if param['canOthers'] == 'true':
                            param['value'].append('Others')
                     for no, value in enumerate(param['value'], start=1):
                            print(str(no) + " : " + str(value))
                     inputParam = input("\nPlease input " + key + " : ")

                     if inputParam.isdigit() and 1 <= int(inputParam) <= len(optionBefore):
                            if param['dataType'] == 'array':
                                   params[key] = []
                                   params[key].append(optionBefore[int(inputParam) - 1])
                                   arrayParam.update({key: inputParam})
                            else:
                                   params[key] = optionBefore[int(inputParam) - 1]
                     else:
                            others = input("1 = To input your " + str(key) + ", 2 = To input your txt file : ")
                            if others == '1':
                                   if param['dataType'] == 'array':
                                          params[key] = []
                                          params[key].append(input("Please input your " + str(key) + " = "))
                                   else:
                                          params[key] = input("Please input your " + str(key) + " = ")
                            elif others == '2':
                                   txtFile = input("Please input your txt file name : ")
                                   delimiter = input("Please input delimiter in your file (it's only if you need more than one param) : ")
                                   try :
                                          file=open("Data/"+txtFile+".txt","r")
                                          data=file.readlines()
                                          file.close()
                                          if param['dataType'] == 'array':
                                                 params[key] = data
                                          else:
                                                 params[key] = data[0]
                                   except:
                                          print ('FILE NOT FOUND OR UNREADABLE IN FOLDER')
                                          print ('CHECK AGAIN !!!')
                                          exit()
                            else:
                                   print ("Wrong Input!")
              else:
                     if param['canOthers'] == 'true':
                            others = input("1 = To input your " + str(key) + ", 2 = To input your txt file : ")
                            if others == '1':
                                   inputParam = input("\nPlease input " + key + " : ")
                                   if param['multiple'] == 'true':
                                          inputParam = inputParam.split(param['delimiter'])
                                          if len(inputParam) > 1:
                                                 for i, value in enumerate(inputParam):
                                                        if i < len(multiParam):
                                                               multiParam[i][key] = value
                                                        else:
                                                               param_dict = {key: value}
                                                               multiParam.append(param_dict)
                                          else:
                                                 if param['dataType'] == 'array':
                                                        arrayParam.update({key: inputParam[0]})
                                                        multiParam.insert(0,arrayParam)
                                                 else:
                                                        params[key] = inputParam[0]
                                   else:
                                          params[key] = inputParam
                            else:
                                   stringParam = ''
                                   noParam = 1
                                   skipConditions = False
                                   needMultipleParam = []
                                   for key, needParam in choosedTestCase['Params'].items():
                                          if "conditions" in needParam and isinstance(needParam['determiningParam'], str):
                                                 if needParam['determiningParam'] in params:
                                                        if "conditions" in needParam and params[param['determiningParam']][0] not in needParam["conditions"]:
                                                               continue  # Skip the current iteration and move to the next one
                                                 else:
                                                        continue
                                          
                                          if "determiningParam" in needParam and isinstance(needParam['determiningParam'], list):
                                                 for determiningParam in needParam['determiningParam']:
                                                        if determiningParam in params:
                                                               if params[determiningParam][0] not in needParam["conditions"][determiningParam]:
                                                                      skipConditions = True
                                                                      break  # Skip the current iteration and move to the next one      
                                                               else:
                                                                      skipConditions = False
                                                                      break
                                                        else:
                                                               skipConditions = True
                                                               continue    
                                          
                                          if skipConditions:
                                                 skipConditions = False
                                                 continue
                                          
                                          if key in params and params[key] != '':
                                                 continue

                                          if stringParam == '':
                                                 stringParam = key
                                          else: 
                                                 stringParam = stringParam+" | "+key
                                          
                                          detailsKey = ''
                                          if "details" in needParam:
                                                 detailsKey = "("+needParam['details']+")"
                                          
                                          strNeedParam = str(noParam)+". "+key+" "+detailsKey+""
                                          needMultipleParam.append(strNeedParam)
                                          noParam = int(noParam)+1

                                   print("This function is need "+str(len(needMultipleParam))+" params :")
                                   
                                   for valParam in needMultipleParam:
                                          print (valParam)

                                   print("You can use this like an example ('|' is using as delimiter) : "+stringParam)
                                   
                                   txtFile = input("Please input your txt file name: ")
                                   delimiter = input("Please input the delimiter in your file: ")

                                   try:
                                          with open("Data" + txtFile + ".txt", "r") as file:
                                                 data = file.readlines()
                                          
                                          if len(data) > 1:
                                                 for listData in data:
                                                        datas = listData.split(delimiter)
                                                        if len(datas) >= 1:
                                                               i = 0
                                                               tempParam = {}
                                                               for key, needParam in choosedTestCase['Params'].items():
                                                                      if "conditions" in needParam and isinstance(needParam['determiningParam'], str):
                                                                             if needParam['determiningParam'] in params:
                                                                                    if "conditions" in needParam and params[param['determiningParam']][0] not in needParam["conditions"]:
                                                                                           continue  # Skip the current iteration and move to the next one
                                                                             else:
                                                                                    continue
                                                                      
                                                                      if "determiningParam" in needParam and isinstance(needParam['determiningParam'], list):
                                                                             for determiningParam in needParam['determiningParam']:
                                                                                    if determiningParam in params:
                                                                                           if params[determiningParam][0] not in needParam["conditions"][determiningParam]:
                                                                                                  skipConditions = True
                                                                                                  break  # Skip the current iteration and move to the next one      
                                                                                           else:
                                                                                                  skipConditions = False
                                                                                                  break
                                                                                    else:
                                                                                           skipConditions = True
                                                                                           continue   
                                                                      if skipConditions:
                                                                             skipConditions = False
                                                                             continue
                                                                      
                                                                      if key in params and params[key] != '':
                                                                             continue

                                                                      param_dict = {key: datas[i]}
                                                                      tempParam.update(param_dict)
                                                                      i = i+1
                                                               multiParam.append(tempParam)                   
                                                        else:
                                                               if param['dataType'] == 'array':
                                                                      arrayParam.update({key: inputParam[0]})
                                                                      multiParam.insert(0,arrayParam)
                                                               else:
                                                                      params[key] = datas[0]
                                          else:
                                                 inputParam = data[0].split(delimiter)
                                                 if len(inputParam) > 1:
                                                        for i, value in enumerate(inputParam):
                                                               if i < len(multiParam):
                                                                      multiParam[i][key] = value
                                                               else:
                                                                      param_dict = {key: value}
                                                                      multiParam.append(param_dict)
                                                 else:
                                                        if param['dataType'] == 'array':
                                                               arrayParam.update({key: inputParam[0]})
                                                               multiParam.insert(0,arrayParam)
                                                        else:
                                                               params[key] = inputParam[0]

                                   except FileNotFoundError:
                                          print("File not found or unreadable in folder.")
                                          print("Please check again!")
                                          exit()

                                   except Exception as e:
                                          print("An error occurred:", str(e))

                                   stopLoop = True
                                   break
                     else:
                            inputParam = input("\nPlease input " + key + " : ")
                            if param['multiple'] == 'true':
                                   inputParam = inputParam.split(param['delimiter'])
                                   if len(inputParam) > 1:
                                          for i, value in enumerate(inputParam):
                                                 if i < len(multiParam):
                                                        multiParam[i][key] = value
                                                 else:
                                                        param_dict = {key: value}
                                                        multiParam.append(param_dict)
                                   else:
                                          if param['dataType'] == 'array':
                                                 if len(multiParam):
                                                        arrayParam.update({key: inputParam[0]})
                                                 else:
                                                        arrayParam.update({key: inputParam[0]})
                                                        multiParam.insert(0,arrayParam)
                                          else:
                                                 params[key] = inputParam[0]
                            else:
                                   params[key] = inputParam

              if stopLoop:
                     break

       if len(multiParam) > 0:
              if params:
                     for item in multiParam:
                            item.update(params)
              function(choosedTestCase['EventName'], multiParam, neededParams)
       else:
              function(choosedTestCase['EventName'], params, neededParams)
else:
    print("\nSorry, the item is not ready\n")



