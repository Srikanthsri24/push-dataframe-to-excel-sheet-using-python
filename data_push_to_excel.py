#with random numbers to create a data frame


# import pandas as pd
# import numpy as np
#
# # Set the random seed for reproducibility
# np.random.seed(42)
#
# # Define the size of the data frame
# num_rows = 5
# num_cols = 3
#
# # Create random data using numpy
# data = np.random.rand(num_rows, num_cols)
#
# # Create a pandas DataFrame
# df = pd.DataFrame(data, columns=['Column1', 'Column2', 'Column3'])
#
# # Print the DataFrame
# print(df)




#with Existed data to create a data frame

# import pandas as pd
#
# # Create a dictionary with the data
# data = {
#     'Column1': [1, 2, 3, 4, 5],
#     'Column2': ['A', 'B', 'C', 'D', 'E'],
#     'Column3': [0.1, 0.2, 0.3, 0.4, 0.5]
# }
#
# # Create a pandas DataFrame
# df = pd.DataFrame(data)
#
# # Print the DataFrame
# print(df)



#to send the data frame to excel sheet using pandas and openpyxl

import pandas as pd
from openpyxl import load_workbook

# Load the Excel file
excel_file_path = 'file path.xlsx'
sheet_name = 'Sheet1'

# Read the existing data from the sheet
existing_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Create a new DataFrame
new_data = pd.DataFrame({'Column1': [0, 2, 4], 'Column2': ['S', 'R', 'I'], 'Column3': [0.0, 0.1, 0.4]})

# Concatenate the existing and new data
concatenated_data = pd.concat([existing_data, new_data], ignore_index=True)

# Write the concatenated data to the Excel file
with pd.ExcelWriter(excel_file_path, mode='openpyxl') as writer:
    writer.book = load_workbook(excel_file_path)
    concatenated_data.to_excel(writer, sheet_name=sheet_name, index=False)

print('Data appended successfully.')
