import pandas as pd
import numpy as np
import re
pd.set_option('display.max_colwidth', None)


# ==============================
#           LOAD po
# ==============================
def loadPO(filename: str):
    """
    This function loads Purchase Order (PO) data from a .csv or .xlsx file and processes it.

    Parameters:
    filename (str): The path to the .csv or .xlsx file to load.

    Returns:
    pd.DataFrame: A DataFrame containing the processed PO data. The DataFrame only includes the 'Type', 'Date', 'Num', 'Source Name', 'Item', 'Qty', 'Cost Price', 'Item Description', 'Override 1', and 'Override 2' columns from the original data. The 'Cost Price' column is rounded to 2 decimal places, the 'Part Number' is extracted from the 'Item' column and trailing whitespace is removed, and the 'Date' column is converted to datetime format.

    Raises:
    Prints an error message and ends execution if the file is not a .csv or .xlsx file.
    """
    if filename.endswith('.csv'):
        po = pd.read_csv(filename)
    elif filename.endswith('.xlsx'):
        po = pd.read_excel(filename)
    else:
        print('ERROR: PO must be .csv or .xlsx file')
        return
    po = po[['Type', 'Date', 'Num', 'Source Name', 'Item', 'Qty', 'Cost Price', 'Item Description', 'Override 1', 'Override 2']] # Only keep relevant columns
    po['Cost Price'] = po['Cost Price'].round(2) # Round Cost Price to 2 decimal places
    po['Part Number'] = po['Item'].str.split('(').str[0] # Extract Part Number from item column
    po['Part Number'] = po['Part Number'].str.strip() # Get rid of trailing whitespace
    po['Date'] = pd.to_datetime(po['Date'], format='%m/%d/%y') # Convert Date to datetime format
    return po

def loadBOM(filename: str):
    """
    This function loads Bill of Materials (BOM) from an Excel file and processes the 'Machined Copy' and 'Purchased Copy' sheets.

    Parameters:
    filename (str): The path to the .xlsx file to load.

    Returns:
    tuple: A tuple containing two DataFrames. The first DataFrame contains the processed 'Machined Copy' sheet and the second DataFrame contains the processed 'Purchased Copy' sheet.

    Raises:
    Prints an error message and ends execution if the file is not a .xlsx file.
    """
    if filename.endswith('.xlsx'):
        po = pd.read_excel(filename)
    else:
        print('ERROR: BOM must be .xlsx file')
        return
    # Machine BOM
    machined = pd.read_excel(filename, sheet_name='Machined Copy')
    machined = machined[['Part #', 'Rev', 'Machined', 'Description', 'Cost', 'Total Qty', 'Vendor', 'Locations']] # Only keep relevant columns
    machined.rename(columns={'Part #': 'Part Number', 'Total Qty': 'Qty'}, inplace=True)
    machined['Part Number'] = machined['Part Number'].astype('str') # Convert Part Number to string
    # Purchased BOM
    purchased = pd.read_excel(filename, sheet_name='Purchased Copy')
    purchased = purchased[['Purchased', 'Description', 'PK QTY', 'Locations']] # Only keep relevant columns
    purchased.rename(columns={'Purchased': 'Part Number', 'PK QTY': 'Pack Qty'}, inplace=True)
    purchased['Part Number'] = purchased['Part Number'].str.replace('\t', '') # get rid of \t tab characters in Part Number  
    return machined, purchased

# ==============================
#           HELPERS
# ==============================

def mapPOtoPurchased(po: pd.DataFrame, purchased: pd.DataFrame):
    po = po.merge(purchased[['Part Number', 'Pack Qty']], on='Part Number', how='left')
    po['Pack Qty'] = po['Pack Qty'].fillna(1) # If Pack Qty is NaN, set to 1
    po['Unit Qty'] = po['Qty'] * po['Pack Qty'] # Multiply Qty by Pack Qty
    po['Unit Price'] = po['Cost Price'] / po['Pack Qty'] # Calculate Unit Price
    po['Part Number'] = po['Part Number'].str.replace('\t', '') # get rid of \t tab characters in Part Number
    # Reorder po columns
    po = po[['Type', 'Date', 'Num', 'Source Name', 'Part Number', 'Item', 'Item Description', 'Cost Price', 'Qty', 'Pack Qty', 'Unit Price', 'Unit Qty', 'Override 1', 'Override 2']]
    return po

def lookupPartNumber(data: pd.DataFrame, part_num: str):
    """
    This function looks up a part number in a given DataFrame.

    Parameters:
    data (pd.DataFrame): The DataFrame in which to look up the part number.
    part_num (str): The part number to look up.

    Returns:
    pd.DataFrame: A DataFrame containing the rows where the part number matches the input part number.
                  If no match is found, a message is printed and None is returned.
    """
    lookup = data[data['Part Number'] == part_num]
    if lookup.empty:
        print('Part Number not found')
    else:
        return lookup

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

def code_locations(location):
    if location is np.nan:
        return np.nan
    elif location.startswith('001'):
        return 'Main Controls'
    elif location.startswith('1'):
        return 'Main Frame'
    elif location.startswith('2') or location.startswith('002'):
        return 'Unwind/Punch Station'
    elif location.startswith('3') or location.startswith('003') or location.startswith('007'):
        return 'Spout Station'
    elif location.startswith('4') or location.startswith('004'):
        return 'Side Seal Station'
    elif location.startswith('5') or location.startswith('005'):
        return 'Cross Seal Station'
    elif location.startswith('6') or location.startswith('006') or location.startswith('008'):
        return 'Cap Station'
    elif location.startswith('8') or location.startswith('7') or location.startswith('009'):
        return 'Delivery/Cutoff Station'
    return np.nan

def get_unique_parts(po: pd.DataFrame, machined: pd.DataFrame, purchased: pd.DataFrame):
    """
    This function gets the unique part numbers from the 'Part Number' columns of three DataFrames: po, machined, and purchased.

    Parameters:
    po (pd.DataFrame): The DataFrame containing Purchase Order data. It must have a 'Part Number' column.
    machined (pd.DataFrame): The DataFrame containing Machined data. It must have a 'Part Number' column.
    purchased (pd.DataFrame): The DataFrame containing Purchased data. It must have a 'Part Number' column.

    Returns:
    np.ndarray: A 1D numpy array containing the unique part numbers from the 'Part Number' columns of the three input DataFrames. The part numbers are sorted in ascending order.
    """
    parts = np.unique(np.concatenate((po['Part Number'].unique().astype(str), machined['Part Number'].unique().astype(str), purchased['Part Number'].unique().astype(str))))
    parts = parts[parts != 'nan']
    return parts
def find_machined_parts(parts: np.ndarray):
    # Match GF12.XXX.XX with an optional .A-Z at the end
    pattern = r'^GF12\.\d{3}\.\d{2}(\.[A-Z])?$'
    return [part for part in parts if re.search(pattern, part)]
# ==============================
#           PROCESS
# ==============================
def process_purchased_part(part_num: str, po: pd.DataFrame, purchased: pd.DataFrame, verbose=True):
    # Filter Data
    purchased_frame = po[po['Part Number'] == part_num].copy() # Get all rows with Part Number
    purchased_frame.sort_values(by=['Date', 'Num'], inplace=True) # Sort by Date, then PO Number
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
    if purchased_frame.empty: # If Part Number is not found in PO, we assume that the part is in stock
        if verbose:
            print('ERROR: Part Number not found in PO')
        final_frame = [] # Create empty list
        for index, row in locations.iterrows(): # Iterate through each location
            final_frame.append(['<Stock>', np.nan, None, np.nan, lookup_frame['Part Number'].iloc[0], lookup_frame['Description'].iloc[0], row['Location'], row['Qty'], 0]) 
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
            final_frame.append(['<Stock>', np.nan, None, np.nan, lookup_frame['Part Number'].iloc[0], lookup_frame['Description'].iloc[0], row['Location'], row['Qty'], 0])
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
def apply_overrides(frame: pd.DataFrame, po: pd.DataFrame):
    po = po[~po['Override 1'].isna()]
    frame = frame[~frame['Num'].isna()].copy()
    frame['Num'] = frame['Num'].astype('int').copy()
    frame = frame.merge(po[['Num', 'Part Number', 'Override 1', 'Override 2']], on=['Num', 'Part Number'], how='left')
    frame['Category 1'] = [override if override is not np.nan else category for override, category in zip(frame['Override 1'], frame['Category 1'])]
    frame['Category 2'] = [override if override is not np.nan else category for override, category in zip(frame['Override 2'], frame['Category 2'])]
    frame.drop(['Override 1', 'Override 2'], axis=1, inplace=True)
    return frame
def process_parts(po: pd.DataFrame, machined: pd.DataFrame, purchased: pd.DataFrame, verbose=False):
    part_nums = get_unique_parts(po, machined, purchased)
    machined_parts = find_machined_parts(part_nums)
    other_parts = part_nums[~np.isin(part_nums, machined_parts)]
    if verbose:
        print(f'Found {len(machined_parts)} machined parts and {len(other_parts)} other parts, with {len(part_nums)} total unique parts.')
        print()
    other_frames = []
    for part in other_parts:
        other_frames.append(process_purchased_part(part, po, purchased, verbose=False))
    other_frames = pd.concat(other_frames)
    other_frames['Category 1'] = other_frames['Location'].apply(code_locations)
    other_frames['Category 2'] = np.nan
    other_frames = apply_overrides(other_frames, po)
    # display(other_frames)
    return other_frames
print("Refreshed.")