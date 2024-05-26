"""
(Purpose) IAC.py is a module that contains functions used in the IAC report
"""

def Utility(filepath:str):
    """
    Extract Data from Energy Charts.xlsx
    Save statistics to Utility.json
    """
    import json, openpyxl
    from easydict import EasyDict
    # Open Utility.json as dictionary
    utility = EasyDict(json.load(open('utility.json')))
    # Read Energy Charts.xlsx
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # Get Raw Data worksheet
    ws = wb['Raw Data']
    ## Save statistics to Utility.json
    # Get Electricity cost from cell D21, 3 digits
    utility.EC.value = round(ws['D21'].value,3)
    # Get Demand cost from cell D23, 2 digits
    utility.DC.value = round(ws['D23'].value,2)
    # Get Fees from cell G19, 2 digits
    Fees = round(ws['G19'].value,2)
    # Get Fuel cost from cell D24, 2 digits
    utility.FC.value = round(ws['D24'].value,2)
    # Get Fuel type from cell Q2, string
    utility.FuelType.value = ws['Q2'].value
    # Get Fuel unit from cell Q3, string
    utility.FuelUnit.value = ws['Q3'].value
    # Get Start Month from cell B7, string
    utility.StartMo.value = ws['B7'].value
    # Get End Month from cell B18, string
    utility.EndMo.value = ws['B18'].value
    # Get Total Electricity kWh from cell C19
    utility.TotalEkWh.value = round(ws['C19'].value)
    # Get Total Electricity MMBtu from cell I19
    utility.TotalEBtu.value = round(ws['I19'].value)
    # Get Total Demand kW from cell E19
    utility.TotalDkW.value = round(ws['E19'].value)
    # Get Total Fuel MMBtu from cell M19
    utility.TotalFBtu.value = round(ws['M19'].value)

    # Get Total Energy worksheet
    ws = wb['Total Energy']
    # Get Total Electricity cost from cell E5+E6
    utility.TotalECost.value = round(ws['E5'].value+ws['E6'].value+Fees)
    # Get Total Fuel cost from cell E7
    utility.TotalFCost.value = round(ws['E7'].value)
    # Get Total Energy MMBtu from cell D8
    utility.TotalBtu.value = round(ws['D8'].value)
    # Get Total Energy Cost from cell E8
    utility.TotalCost.value = round(ws['E8'].value)

    # Write Natural Gas Cost for compatibility.
    if utility.FuelType.value == "Natural Gas":
        utility.NGC.value = utility.FC.value
    else:
        utility.NGC.value = 0

    # Save to Utility.json
    with open('Utility.json', 'w') as f:
        json.dump(utility, f, indent=4)

                
def title_case(text: str) -> str:
    """
    Make title case in natural language
    :param text: String
    :return: formatted string with title case in natural language
    """
    lowerexceptions = ["a","an","and","as","at","but","by","for","from","if","in","nor","not","of","off","on","or","per","the","to","so","up","via","with","yet"]
    upperexceptions = ["VFD","(VFD)","LED","AC","HVAC"]
    text = text.split()
    # Capitalize every word that is not on "exceptions" list
    for i, word in enumerate(text):
        if i==0:
            text[i] = word.title()
        elif word.lower() in lowerexceptions:
            text[i] = word.lower()
        elif word.upper() in upperexceptions:
            text[i] = word.upper()
        else:
            text[i] = word.title()
    return ' '.join(text)


def validate_arc(ARC):
    """
    Validate ARC number
    :param ARC: Full ARC as a string
    """
    import os, json
    # Validate if ARC is in x.xxxx.xxx format
    ARCsplit = ARC.split('.')
    if len(ARCsplit) != 3:
        raise Exception("ARC number must be in x.xxx(x).x format")
    # if ARC split are nut full numbers
    for i in range(len(ARCsplit)):
        if ARCsplit[i].isdigit() == False:
            raise Exception("ARC number must be in x.xxx(x).x format")
    
    # Parse ARC code
    code = ARCsplit[0] + '.' + ARCsplit[1]
    # Read ARC.json as dictionary
    arc_path = os.path.dirname(os.path.abspath(__file__))
    ARCdict = json.load(open(os.path.join(arc_path, 'ARC.json')))
    try:
        desc = ARCdict[code]
        print(code + ": "+ desc)
    except:
        raise Exception("ARC not found.") 

    # Parse application code
    app = ARCsplit[2]
    if app == '1':
        print("Application code 1: Manufacturing Process")
    elif app == '2':
        print("Application code 2: Process Support")
    elif app == '3':
        print("Application code 3: Building and Grounds")
    elif app == '4':
        print("Application code 4: Administrative")
    else:
        raise Exception("Application code not found.")

    print("")


def grouping_num(dic: dict) -> dict:
    """
    Add thousand separator to numbers in a dictionary and format it to string
    :param dic: EasyDict
    :return: Dictionary with keys in thousand separator
    """
    import locale, numpy
    # set locale to US
    locale.setlocale(locale.LC_ALL, 'en_US')
    for key in dic.keys():
        if type(dic[key]) == int or type(dic[key]) == numpy.int64:
            dic[key] = locale.format_string('%d', dic[key], grouping=True)
        elif type(dic[key]) == float or type(dic[key]) == numpy.float64:
            dic[key] = locale.format_string('%g', dic[key], grouping=True)
        # if dic[key] is a ndarray
        elif type(dic[key]) == numpy.ndarray:
            dic[key] = dic[key].tolist()
            for i in range(len(dic[key])):
                if type(dic[key][i]) == int:
                    dic[key][i] = locale.format_string('%d', dic[key][i], grouping=True)
                elif type(dic[key][i]) == float:
                    dic[key][i] = locale.format_string('%g', dic[key][i], grouping=True)
        else:
            pass
    return dic

def dollar(varlist: list, dic: dict, digits: int=0) -> str:
    """
    Format numbers in a dictionary and to currency string
    :param varlist: List of keys in the dictionary
    :param dic: EasyDict
    :param digits: Number of digits, default is 0
    :return: Dictionary with keys in formatted currency string
    """
    import locale
    # if varlist is not a list of strings
    if type(varlist) != list:
        raise Exception("Variable list must be a list of strings")
    for var in varlist:
        if type(var) != str:
            raise Exception("Variable list must be a list of strings")
        if var not in dic.keys():
            raise Exception("Variable not found in dictionary")
    # if digits is not a natural number
    if type(digits) != int:
        raise Exception("Digits must be a natural number")
    if digits < 0:
        raise Exception("Digits must be a natural number")
    # set locale to US
    locale.setlocale(locale.LC_ALL, 'en_US')
    locale._override_localeconv={'frac_digits':digits}
    for var in varlist:
        dic[var] = locale.currency(dic[var], grouping=True)
    return dic


def combine_words(words: list) -> str:
    """
    :param words: list of strings
    :return: string of words separated by "," and "and"
    """
    # if words is not a list
    if type(words) != list:
        raise Exception("Input must be a list of strings")
    combined = ""
    for i in range(len(words)):  
        # if word is not a string
        if type(words[i]) != str:
            raise Exception("Input must be a list of strings")
        combined = combined + words[i]    
        if i < len(words) - 2:
            combined = combined + ', '
        if i == len(words) - 2:
            combined = combined + ' and ' 
        else:
            pass
    return combined


def add_image(doc, tag: str, image_path: str, wd):
    """
    Add image to Word document, search for tag in doc and replace with the image
    :param doc: Document
    :param tag: Image tag as string
    :param image_path: Path to the image as string
    :param wd: Image width
    :return: None
    """
    import os
    # if tag is not a string
    if type(tag) != str:
        raise Exception("Tag must be a string")
    # if image file is not found
    if os.path.isfile(image_path) == False:
        raise Exception("Image file not found")
    found_tag = False
    for p in doc.paragraphs:
        if tag in p.text:
            p.text = p.text.replace(tag, '')
            r = p.add_run()
            r.add_picture(image_path, width=wd)
            found_tag = True
            break
    if found_tag == False:
        # Throw error if tag is not found 
        raise Exception("Tag "+ tag +" not found")


def payback(ACS, IC) -> str:
    """
    Format payback period by year and month
    :param ACS: Annual Cost Savings ($/yr)
    :param IC: Implementation Cost ($)
    :return: formatted Payback Period as string
    """
    import math, numpy
    # if the dtype is numpy
    if type(ACS) == numpy.int64 or type(ACS) == numpy.float64:
        ACS = ACS.item()
    if type(IC) == numpy.int64 or type(IC) == numpy.float64:
        IC = IC.item()
    # if ACS or IC is not a number
    if type(ACS) != int and type(ACS) != float:
        raise Exception("Annual Cost Savings must be a number")
    if type(IC) != int and type(IC) != float:
        raise Exception("Implementation Cost must be a number")
    # Immediate payback
    if IC == 0:
        return "Immediate"
    # Infinite or negative payback
    if ACS <= 0:
        return "Infinite"
    PB = IC / ACS
    if PB <= 11.0 / 12.0:
        PB = math.ceil(PB * 12.0)
        PBstr = str(PB) + " month"
    else:
        PB = math.ceil(PB * 10.0) / 10.0
        PBstr = str(PB).rstrip("0").rstrip(".") + " year"
    if PB > 1.0:
        PBstr = PBstr + "s"
    return PBstr