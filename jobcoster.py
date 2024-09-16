import pandas as pd
import numpy as np
import re
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

    machined = machined[['Part #', 'Rev', 'Machined', 'Description', 'Cost', 'Total Qty', 'Mfg/Matl', 'Vendor', 'Locations']]
    machined.rename(columns={'Part #': 'Part Number', 'Total Qty': 'Total QTY'}, inplace=True)
    # remove rows where Part Number is NaN or ' '
    machined = machined[~machined['Part Number'].isnull()]
    machined = machined[machined['Part Number'] != ' ']
    
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

def get_unique_parts(po: pd.DataFrame, machined: pd.DataFrame, purchased: pd.DataFrame):
    parts = np.unique(np.concatenate((po['Part Number'].unique().astype(str), machined['Part Number'].unique().astype(str), purchased['Part Number'].unique().astype(str))))
    parts = parts[parts != 'nan']
    return parts

def find_machined_parts(parts: np.ndarray):
    # Match GF12.XXX.XX with an optional .A-Z at the end
    pattern = r'^GF12\.\d{3}\.\d{2}(\.[A-Z])?$'
    return [part for part in parts if re.search(pattern, part)]
def parse_locations(locations: str):
    """
    This function parses a string of locations and quantities, and returns a DataFrame with separate columns for location and quantity.

    Parameters:
    locations (str): A string of locations and quantities, separated by commas. Each location and quantity is in the format 'locationxquantity'.

    Returns:
    pd.DataFrame: A DataFrame with two columns: 'Location' and 'Qty'. The 'Location' column contains the location strings, and the 'Qty' column contains the quantities as floats rounded to 2 decimal places.
    """
    locations = locations.replace(' ', '') # Remove whitespace
    locations = [location for location in locations.split(',') if location != ''] # Split by comma and remove empty strings (i.e starting and ending commas)
    
    # Split apart location and quantity
    location_list = []
    for location in locations:
        info = location.split('x')
        info.reverse()
        location_list.append(info)
    return pd.DataFrame(location_list, columns=['Location', 'Qty']).astype({'Qty': 'float'}).round({'Qty': 2})

def process_purchased_part(part_num: str, po: pd.DataFrame, purchased: pd.DataFrame, verbose=True):
    # Filter Data
    purchased_frame = po[po['Part Number'] == part_num].copy() # Get all rows with Part Number
    purchased_frame.sort_values(by=['Date', 'PO #'], inplace=True) # Sort by Date, then PO Number
    lookup_frame = purchased[purchased['Part Number'] == part_num].copy() # Get all rows with Part Number

    # Reset indexes
    purchased_frame.reset_index(drop=True, inplace=True)
    lookup_frame.reset_index(drop=True, inplace=True)

    # Get Locations
    try:
        locations = parse_locations(lookup_frame['Locations'].iloc[0])
        # locations.loc[0, 'Qty'] = 10
        if verbose:
            print(f'Part Number: {part_num}')
            display(locations)
    except Exception:
        pass
    # Process Data
    purchased_frame['Type'] = "Purchase Order"
    purchased_frame.rename(columns={'Order QTY': 'Unit Qty', 'PO Cost': 'Unit Price'}, inplace=True)
    if purchased_frame.empty: # If Part Number is not found in PO, we assume that the part is in stock
        if verbose:
            print('ERROR: Part Number not found in PO')
        final_frame = [] # Create empty list
        for index, row in locations.iterrows(): # Iterate through each location
            final_frame.append(['<Stock>', np.nan, None, np.nan, lookup_frame['Part Number'].iloc[0], lookup_frame['Description'].iloc[0], row['Location'], row['Qty'], 0]) 
        final_frame = pd.DataFrame(final_frame, columns=['Type', 'Date', 'PO #', 'Vendor', 'Part Number', 'Description', 'Location', 'Unit Qty', 'Unit Price'])
        return final_frame
    if lookup_frame.empty: # If Part Number is not found in the Purchased BOM, this would imply the part was purchased after the Quote
        if verbose:
            print('ERROR: Part Number not found in Purchased BOM')
        final_frame = purchased_frame[['Type', 'Date', 'PO #', 'Vendor', 'Part Number', 'Description', 'Unit Qty', 'Unit Price']]
        final_frame.insert(6, 'Location', np.nan)
        final_frame.reset_index(drop=True, inplace=True)
        return final_frame
    # OTHERWISE, we have a match in both PO and Purchased BOM
    final_frame = [] # Create empty list
    for index, row in locations.iterrows(): # Iterate through each location
        # No Qty match, this implies that the remaining Qty is in stock
        if purchased_frame.empty:
            final_frame.append(['<Stock>', np.nan, None, np.nan, lookup_frame['Part Number'].iloc[0], lookup_frame['Description'].iloc[0], row['Location'], row['Qty'], 0])
        # Exact Qty match
        elif row['Qty'] == purchased_frame['Unit Qty'].iloc[0]:
            if verbose:
                print('Exact Qty match')
            final_frame.append(purchased_frame[['Type', 'Date', 'PO #', 'Vendor', 'Part Number', 'Description']].iloc[0].to_list() + [row['Location'], row['Qty'], purchased_frame['Unit Price'].iloc[0]])
            purchased_frame.drop(0, inplace=True) # Drop the Qty row from PO
            purchased_frame.reset_index(drop=True, inplace=True)
        # Qty in Purchased is less than Qty in PO
        elif row['Qty'] < purchased_frame['Unit Qty'].iloc[0]:
            if verbose:
                print('Qty in Purchased is less than Qty in PO')
            final_frame.append(purchased_frame[['Type', 'Date', 'PO #', 'Vendor', 'Part Number', 'Description']].iloc[0].to_list() + [row['Location'], row['Qty'], purchased_frame['Unit Price'].iloc[0]])
            purchased_frame.loc[0, 'Unit Qty'] -= row['Qty'] # Subtract Qty from PO
    # Deal with extra Qty in PO
    if not purchased_frame.empty: # If PO is not empty, this implies that there is extra Qty that was not accounted for in the Purchased BOM
        if verbose:
            print('Qty in Purchased is greater than Qty in PO')
        for index, row in purchased_frame.iterrows():
            final_frame.append(purchased_frame[['Type', 'Date', 'PO #', 'Vendor', 'Part Number', 'Description']].iloc[index].to_list() + ['<Extra>', row['Unit Qty'], row['Unit Price']])
    # if verbose:
        # display(purchased_frame)
    # Convert final_frame to DataFrame and format columns
    final_frame = pd.DataFrame(final_frame, columns=['Type', 'Date', 'PO #', 'Vendor', 'Part Number', 'Description', 'Location', 'Unit Qty', 'Unit Price'])
    return final_frame

def processPurchasedParts(po: pd.DataFrame, purchased: pd.DataFrame, machined: pd.DataFrame, extrusion: pd.DataFrame, bolts: pd.DataFrame, assemblies: pd.DataFrame, verbose: bool = True):
    print("Processing Purchased Parts...")
    part_nums = get_unique_parts(po, machined, purchased)
    machined_parts = find_machined_parts(part_nums)
    other_parts = part_nums[~np.isin(part_nums, machined_parts)]
    if verbose:
        print(f'Found {len(machined_parts)} machined parts and {len(other_parts)} other parts, with {len(part_nums)} total unique parts.')
        print()
    machined_frames = []
    machined_parts = [re.sub(r'\.[A-Z]$', '', part) for part in machined_parts]
    machined_parts = np.unique(machined_parts)

    other_frames = []
    for part in other_parts:
        other_frames.append(process_purchased_part(part, po, purchased, verbose=False))
    other_frames = pd.concat(other_frames)
    # other_frames['Category 1'] = other_frames['Location'].apply(code_locations)
    # other_frames['Category 2'] = np.nan
    # other_frames = apply_overrides(other_frames, po)
    # display(other_frames)
    return other_frames









print("jobcoster.py loaded successfully!")