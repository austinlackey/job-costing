import pandas as pd

import matplotlib.pyplot as plt

# Sample dataframe
df = pd.DataFrame({'A': [1, 2, 3], 'B': [[4, 5], [5, 6], [6, 7, 8, 9, 10, 11, 12]]})

def create_cost_dataframe(df):
    cost_list = []
    
    # Iterate through each row index
    for i in range(len(df)):
        # Access the elements in column B
        for element in df.loc[i, 'B']:
            cost_list.append(element)
    
    # Convert the list to a new dataframe
    new_df = pd.DataFrame(cost_list, columns=['Cost'])
    return new_df

# Create the new dataframe
new_cost_df = create_cost_dataframe(df)

# Plotting the histogram with red bars
plt.hist(new_cost_df['Cost'], bins=range(min(new_cost_df['Cost']), max(new_cost_df['Cost']) + 1), edgecolor='black', color='red')
plt.xlabel('Cost')
plt.ylabel('Frequency')
plt.title('Histogram of Costs')
plt.show()
