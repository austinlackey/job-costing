import pandas as pd
import numpy as np
import os
import re
pd.set_option('display.max_colwidth', None)


# ==============================
#           LOAD DATA
# ==============================
def loadPO(filename: str):
    if filename.endswith('.csv'):
        data = pd.read_csv(filename)
    if filename.endswith('.xlsx'):
        data = pd.read_excel(filename)
    
    data = data[['Type', 'Date', 'Num', 'Source Name', 'Item', 'Qty', 'Cost Price', 'Item Description', 'Override 1', 'Override 2']] # Only keep relevant columns
    data['Cost Price'] = data['Cost Price'].round(2) # Round Cost Price to 2 decimal places
    data['Part Number'] = data['Item'].str.split('(').str[0] # Extract Part Number from item column
    data['Part Number'] = data['Part Number'].str.strip() # Get rid of trailing whitespace
    data['Date'] = pd.to_datetime(data['Date'], format='%m/%d/%y')
    return data

def loadBOMs(filename: str):
    # Check if file exists
    if not os.path.exists(filename):
        print(f"ERROR: File {filename} does not exist.")
        return

    # Check if file is of correct type
    if not filename.endswith('.xlsx'):
        print('ERROR: Only .xlsx files are supported.')
        return

    # Read in excel and get all sheets
    xls = pd.ExcelFile(filename)
    sheet_names = xls.sheet_names
    print(sheet_names)

    # Dictionary to hold dataframes
    df_dict = {}

    # Function mapping for formatting
    format_functions = {
        'assemblies': formatBOM_assemblies,
        'machined': formatBOM_machined,
        'purchased': formatBOM_purchased,
        'extrusion': formatBOM_extrusion
    }

    # Process each sheet
    for sheet in sheet_names:
        for key in format_functions.keys():
            if key in sheet.lower():
                df = pd.read_excel(filename, sheet_name=sheet)
                df = format_functions[key](df)
                df_dict[key] = df

    # Return dictionary of dataframes
    return df_dict


# ==============================
#           FORMAT DATA
# ==============================
def formatBOM_assemblies(data: pd.DataFrame):
    return data
def formatBOM_machined(data: pd.DataFrame):
    try:
        data = data[['Part #', 'Rev', 'Machined', 'Description', 'Cost', 'Total Qty', 'Vendor', 'Locations']].copy() # Only keep relevant columns
        data.rename(columns={'Part #': 'Part Number', 'Machined': 'Part Number Rev'}, inplace=True) # Rename columns
        return data
    except Exception as e:
        print(f"An error occurred: {e}")
        return None
def formatBOM_purchased(data: pd.DataFrame):
    try:
        data = data[['Purchased', 'Description', 'PK QTY', 'Locations', 'Vendor']].copy() # Only keep relevant columns
        data.rename(columns={'Purchased': 'Part Number', 'PK QTY': 'Pack Qty'}, inplace=True) # Rename columns
        data['Part Number'] = data['Part Number'].str.replace('\t', '') # get rid of \t tab characters in Part Number
        data = data[~data['Vendor'].str.contains('crave', case=False)].reset_index(drop=True)
        # get rid of vendor column
        data = data[['Part Number', 'Description', 'Pack Qty', 'Locations']].copy()
        return data
    except Exception as e:
        print(f"An error occurred: {e}")
        return None
def formatBOM_extrusion(data: pd.DataFrame):
    return data

def mapPOtoPurchased(po: pd.DataFrame, purchased: pd.DataFrame):
    po = po.merge(purchased[['Part Number', 'Pack Qty']], on='Part Number', how='left')
    po['Pack Qty'] = po['Pack Qty'].fillna(1) # If Pack Qty is NaN, set to 1
    po['Unit Qty'] = po['Qty'] * po['Pack Qty'] # Multiply Qty by Pack Qty
    po['Unit Price'] = po['Cost Price'] / po['Pack Qty'] # Calculate Unit Price
    po['Part Number'] = po['Part Number'].str.replace('\t', '') # get rid of \t tab characters in Part Number
    # Reorder po columns
    po = po[['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item', 'Item Description', 'Cost Price', 'Qty', 'Pack Qty', 'Unit Price', 'Unit Qty', 'Override 1', 'Override 2']]
    return po

# ==============================
#           HELPERS
# ==============================

def lookupPartNumber(data: pd.DataFrame, part_num: str, verbatim=True):
    lookup = None
    if verbatim:
        lookup = data[data['Part Number'] == part_num]
    else:
        lookup = data[data['Part Number'].str.contains(part_num, case=False)]
    if lookup.empty:
        print('Part Number not found')
    return lookup

def get_unique_part_nums():
    po_part_nums = po['Part Number'].unique()
    purchased_part_nums = purchased['Part Number'].unique()
    part_nums = np.unique(np.concatenate((po_part_nums, purchased_part_nums), axis=0))
    return part_nums

def code_locations(location):
    if pd.isna(location):
        return np.nan

    location_mapping = {
        '001': 'Main Controls',
        '1': 'Main Frame',
        '2': 'Unwind/Punch Station',
        '002': 'Unwind/Punch Station',
        '3': 'Spout Station',
        '003': 'Spout Station',
        '007': 'Spout Station',
        '4': 'Side Seal Station',
        '004': 'Side Seal Station',
        '5': 'Cross Seal Station',
        '005': 'Cross Seal Station',
        '6': 'Cap Station',
        '006': 'Cap Station',
        '008': 'Cap Station',
        '8': 'Delivery/Cutoff Station',
        '7': 'Delivery/Cutoff Station',
        '009': 'Delivery/Cutoff Station'
    }

    for code, name in location_mapping.items():
        if location.startswith(code):
            return name

    return np.nan

def parse_locations(locations):
    locations = locations.replace(' ', '') # Remove whitespace
    locations = [location for location in locations.split(',') if location != ''] # Split by comma and remove empty strings (i.e starting and ending commas)
    
    # Split apart location and quantity
    location_list = []
    for location in locations:
        info = location.split('x')
        info.reverse()
        location_list.append(info)
    return pd.DataFrame(location_list, columns=['Location', 'Qty']).astype({'Qty': 'float'}).round({'Qty': 2})


# ==============================
#           PROCESS
# ==============================

def find_machined_parts(parts: np.ndarray):
    # Match GF12.XXX.XX with an optional .A-Z at the end
    pattern = r'^GF12\.\d{3}\.\d{2}(\.[A-Z])?$'
    return [part for part in parts if re.search(pattern, part)]

def chop_revision(parts: np.ndarray):
    # Match GF12.XXX.XX.[A-Z] and remove the .[A-Z] at the end
    pattern = r'^GF12\.\d{3}\.\d{2}\.[A-Z]$'
    return [part[:-2] if re.search(pattern, part) else part for part in parts]

def process_parts(po: pd.DataFrame, bom_assemblies: pd.DataFrame = None, bom_machined: pd.DataFrame = None, bom_purchased: pd.DataFrame = None, bom_extrusion: pd.DataFrame = None, verbose=False):
    # PO Part Numbers
    po_part_nums = np.array(po['Part Number'].dropna(), dtype=str)

    # Unique Merged Machine Part Numbers
    po_machined_part_nums = find_machined_parts(po_part_nums)
    po_machined_part_nums = chop_revision(po_machined_part_nums)
    unique_machine_part_nums = np.unique(np.concatenate((po_machined_part_nums, bom_machined['Part Number'].dropna().astype(str)), axis=0))
    # print(unique_machine_part_nums)
    output_machined = []
    process_machined_part('GF12.414.02', bom_machined, po, verbose=True)
    # for part_num in unique_machine_part_nums:
    #     output_machined.append(process_machined_part(part_num, bom_machined, po, verbose=False))
    # # Remove empty or all-NA entries
    # output_machined = [df for df in output_machined if df.dropna().shape[0] > 0]
    # output_machined = pd.concat(output_machined, ignore_index=True)
    # print(output_machined)




    # # Unique Merged Purchased Part Numbers
    # unique_purchased_part_nums = np.unique(np.concatenate((po_part_nums, bom_purchased['Part Number'].dropna().astype(str)), axis=0))
    # unique_purchased_part_nums = np.setdiff1d(unique_purchased_part_nums, unique_machine_part_nums)
    # output_purchased = []
    # for part_num in unique_purchased_part_nums:
    #     output_purchased.append(process_purchased_part(part_num, bom_purchased, po, verbose=False))
    # # Remove empty or all-NA entries
    # output_purchased = [df for df in output_purchased if df.dropna().shape[0] > 0]
    # output_purchased = pd.concat(output_purchased, ignore_index=True)
    # # Deal with Freight and Expedited Attributes
    # output_purchased.loc[(output_purchased['Part Number'].fillna('').str.contains('freight|expedit', case=False)) & (output_purchased['Location'].isna()), 'Location'] = 'Freight'
    # # find any rows where Part Number matches unique_machine_part_nums

def process_machined_part(part_num: str, bom_machined: pd.DataFrame, po: pd.DataFrame, verbose=False):
    # Filter Data
    purchased_frame = po[po['Part Number'].str.startswith(part_num)]
    purchased_frame = purchased_frame.sort_values(by=['Date', 'Num']) # Sort by Date, then PO Number
    lookup_frame = bom_machined[bom_machined['Part Number'] == part_num].copy() # Get all rows with Part Number

    print("PO")
    display(purchased_frame)
    print("BOM")
    display(lookup_frame)

def process_purchased_part(part_num: str, bom_purchased: pd.DataFrame, po: pd.DataFrame, verbose=False):
    # Filter Data
    purchased_frame = po[po['Part Number'] == part_num].copy() # Get all rows with Part Number
    purchased_frame.sort_values(by=['Date', 'Num'], inplace=True) # Sort by Date, then PO Number
    lookup_frame = bom_purchased[bom_purchased['Part Number'] == part_num].copy() # Get all rows with Part Number

    # Reset indexes
    purchased_frame.reset_index(drop=True, inplace=True)
    lookup_frame.reset_index(drop=True, inplace=True)

    # Get Locations
    try:
        locations = parse_locations(lookup_frame['Locations'].iloc[0])
        # locations.loc[0, 'Qty'] = 10
        if verbose:
            display(locations)
    except Exception:
        pass

    # Display/Print Data before processing
    if verbose:
        print("Purchase Order")
        display(purchased_frame)
        print("Purchased Tab")
        display(lookup_frame)
    
    # Process Data
    if purchased_frame.empty: # If Part Number is not found in PO, we assume that the part is in stock
        if verbose:
            print('ERROR: Part Number not found in PO')
        final_frame = [] # Create empty list
        for index, row in locations.iterrows(): # Iterate through each location
            final_frame.append(['<Stock>', np.nan, np.nan, np.nan, lookup_frame['Part Number'].iloc[0], lookup_frame['Description'].iloc[0], row['Location'], row['Qty'], 0]) 
        final_frame = pd.DataFrame(final_frame, columns=['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item Description', 'Location', 'Unit Qty', 'Unit Price'])
        return final_frame
    if lookup_frame.empty: # If Part Number is not found in the Purchased BOM, this would imply the part was purchased after the Quote
        if verbose:
            print('ERROR: Part Number not found in Purchased BOM')
        final_frame = purchased_frame[['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item Description', 'Unit Qty', 'Unit Price']]
        final_frame.insert(6, 'Location', np.nan)
        final_frame.reset_index(drop=True, inplace=True)
        return final_frame
    # OTHERWISE, we have a match in both PO and Purchased BOM
    final_frame = [] # Create empty list
    for index, row in locations.iterrows(): # Iterate through each location
        # No Qty match, this implies that the remaining Qty is in stock
        if purchased_frame.empty:
            final_frame.append(['<Stock>', np.nan, np.nan, np.nan, lookup_frame['Part Number'].iloc[0], lookup_frame['Description'].iloc[0], row['Location'], row['Qty'], 0])
        # Exact Qty match
        elif row['Qty'] == purchased_frame['Unit Qty'].iloc[0]:
            if verbose:
                print('Exact Qty match')
            final_frame.append(purchased_frame[['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item Description']].iloc[0].to_list() + [row['Location'], row['Qty'], purchased_frame['Unit Price'].iloc[0]])
            purchased_frame.drop(0, inplace=True) # Drop the Qty row from PO
            purchased_frame.reset_index(drop=True, inplace=True)
        # Qty in Purchased is less than Qty in PO
        elif row['Qty'] < purchased_frame['Unit Qty'].iloc[0]:
            if verbose:
                print('Qty in Purchased is less than Qty in PO')
            final_frame.append(purchased_frame[['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item Description']].iloc[0].to_list() + [row['Location'], row['Qty'], purchased_frame['Unit Price'].iloc[0]])
            purchased_frame.loc[0, 'Unit Qty'] -= row['Qty'] # Subtract Qty from PO
    # Deal with extra Qty in PO
    if not purchased_frame.empty: # If PO is not empty, this implies that there is extra Qty that was not accounted for in the Purchased BOM
        if verbose:
            print('Qty in Purchased is greater than Qty in PO')
        for index, row in purchased_frame.iterrows():
            final_frame.append(purchased_frame[['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item Description']].iloc[index].to_list() + ['<Extra>', row['Unit Qty'], row['Unit Price']])
    if verbose:
        display(purchased_frame)
    # Convert final_frame to DataFrame and format columns
    final_frame = pd.DataFrame(final_frame, columns=['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item Description', 'Location', 'Unit Qty', 'Unit Price'])
    return final_frame