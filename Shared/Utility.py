def Utility(filepath:str):
    """
    Extract Data from Energy Charts.xlsx
    Save statistics to Utility.json
    """
    import os, json, openpyxl
    from easydict import EasyDict
    # Open Utility.json as dictionary
    utility = EasyDict(json.load(open(os.path.join("Shared", "Utility.json"))))
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
    with open(os.path.join("Shared", "Utility.json"), 'w') as f:
        json.dump(utility, f, indent=4)