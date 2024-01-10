import pandas as pd
import numpy as np
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

def loadBOM_Purchased(filename: str):
    if filename.endswith('.csv'):
        data = pd.read_csv(filename)
    if filename.endswith('.xlsx'):
        data = pd.read_excel(filename)
    data = data[['Purchased', 'Description', 'PK QTY', 'Locations']] # Only keep relevant columns
    data.rename(columns={'Purchased': 'Part Number', 'PK QTY': 'Pack Qty'}, inplace=True)
    data['Part Number'] = data['Part Number'].str.replace('\t', '') # get rid of \t tab characters in Part Number  

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

def lookupPartNumber(data: pd.DataFrame, part_num: str):
    lookup = data[data['Part Number'] == part_num]
    if lookup.empty:
        print('Part Number not found')
    else:
        return lookup

def get_unique_part_nums():
    po_part_nums = po['Part Number'].unique()
    purchased_part_nums = purchased['Part Number'].unique()
    part_nums = np.unique(np.concatenate((po_part_nums, purchased_part_nums), axis=0))
    return part_nums

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