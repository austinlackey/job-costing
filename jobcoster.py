import pandas as pd
import numpy as np
pd.set_option('display.max_colwidth', None)

def loadData(filename: str):
    if not filename.endswith('.xlsx'):
        raise ValueError('File must be of type .xlsx')
    # Check that each column is present
    required_columns = ['All Purchase Orders', 'BOM Assemblies', 'BOM Machined', 'BOM Purchased', 'BOM Extrusion', 'BOM Bolts']
    with pd.ExcelFile(filename) as xls:
        for sheet in required_columns:
            if sheet not in xls.sheet_names:
                raise ValueError(f'{sheet} sheet not found in {filename}, either missing or misnamed')
    print("Data loaded successfully!")
    # Read each sheet
    po = pd.read_excel(filename, sheet_name='All Purchase Orders')
    assemblies = pd.read_excel(filename, sheet_name='BOM Assemblies')
    machined = pd.read_excel(filename, sheet_name='BOM Machined')
    purchased = pd.read_excel(filename, sheet_name='BOM Purchased')
    extrusion = pd.read_excel(filename, sheet_name='BOM Extrusion')
    bolts = pd.read_excel(filename, sheet_name='BOM Bolts')

    # Clean up data
    po = po.iloc[2:] # Drop first 2 rows since they are blank
    po = po[['Date', 'Num', 'Source Name', 'Item', 'Qty', 'Amount', 'Item Description']]
    po['Amount'] = po['Amount'].round(2)
    po['Item'] = po['Item'].str.split('(').str[0]
    po.rename(columns={'Num': 'PO #', 'Source Name': 'Vendor', 'Qty': 'Order QTY', 'Item': 'Part Number', 'Item': 'Part Number', 'Item Description': 'Description', 'Amount': 'PO Cost'}, inplace=True)
    # convert PO # to int
    po['PO #'] = po['PO #'].fillna(0).astype(int).astype(str)
    po['Part Number'] = po['Part Number'].str.strip() # Get rid of trailing whitespace
    po = po.reset_index(drop=True)

    assemblies = assemblies[['Job #', 'Assy', 'Item', 'Rev', 'Assembly', 'Description', 'Total Qty', 'Locations']]
    assemblies.rename(columns={'Total Qty': 'Total QTY'}, inplace=True)

    purchased = purchased[['Purchased', 'Description', 'Cost', 'PK QTY', 'BOM QTY', 'Order QTY', 'Vendor', 'Order Date', 'PO #', 'Locations']]   
    purchased.rename(columns={'Purchased': 'Part Number'}, inplace=True)
    purchased['Part Number'] = purchased['Part Number'].str.strip() # Get rid of trailing whitespace
    
    return po, assemblies, machined, purchased, extrusion, bolts

def mergeData(po: pd.DataFrame, assemblies: pd.DataFrame, machined: pd.DataFrame, purchased: pd.DataFrame, extrusion: pd.DataFrame, bolts: pd.DataFrame):
    # Merge PO with Purchased
    po = po.merge(purchased[['Part Number', 'PK QTY', 'Locations']], on='Part Number', how='left')
    po['PK QTY'] = po['PK QTY'].fillna(1) # If Pack Qty is NaN, set to 1
    po['Unit QTY'] = po['Order QTY'] * po['PK QTY'] # Multiply Qty by Pack Qty
    po['Unit Price'] = po['PO Cost'] / po['Unit QTY'] # Calculate Unit Price
    po['Unit Price'] = po['Unit Price'].round(2)
    po['Part Number'] = po['Part Number'].str.replace('\t', '') # get rid of \t tab characters in Part Number
    # po = po[~po['Locations'].isnull()]
    return po

def parseLocations(locations):
    locations = locations.replace(' ', '') # Remove whitespace
    locations = [location for location in locations.split(',') if location != ''] # Split by comma and remove empty strings (i.e starting and ending commas)
    
    # Split apart location and quantity
    location_list = []
    for location in locations:
        info = location.split('x')
        info.reverse()
        location_list.append(info)
    return pd.DataFrame(location_list, columns=['Location', 'Qty']).astype({'Qty': 'float'}).round({'Qty': 2})

def processPurchasedParts(data: pd.DataFrame, verbose=False):
    unique_parts = data['Part Number'].unique()
    final_frame = []
    for part_num in unique_parts:
        # Filter Data
        filtered_data = data[data['Part Number'] == part_num].copy()
        # Sort by Date
        filtered_data = filtered_data.sort_values(by='Date')

        locations = parseLocations(filtered_data['Locations'].iloc[0])
        print(locations)
        display(filtered_data)
    

print("jobcoster.py loaded successfully!")