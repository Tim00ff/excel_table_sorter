import os
import pandas as pd

# Get the current directory
current_directory = os.getcwd()
data_folder = 'DATA'  # Name of the folder containing the data files

# Construct the path to the data folder within the current directory
folder_path = os.path.join(current_directory, data_folder)
output_combined_file = 'combined_data.xlsx'  # Output Excel file

# Prompt the user for the column index and value to sort for
column_index_to_sort = int(input("Enter the index of the column to sort for (0-based index): "))
value_to_sort = input(f"Enter the value in column {column_index_to_sort} to sort for: ")

# Create an empty DataFrame to store the combined data
combined_df = pd.DataFrame()

# Iterate through each file in the data folder
for filename in os.listdir(folder_path):
    file_path = os.path.join(folder_path, filename)
    if os.path.isfile(file_path):  # Check if the item is a file
        print(f"Detected file: {filename}")  # Print the detected filename

        input_df = pd.read_excel(file_path)
        
        # Create a DataFrame with the filename in column A using iloc
        filename_df = pd.DataFrame({input_df.columns[0]: [filename]})
        
        # Filter the rows based on user input for column index and value using iloc
        filtered_df = input_df[input_df.iloc[:, column_index_to_sort] == value_to_sort]
        
        # Concatenate the filename DataFrame and the filtered data
        combined_df = pd.concat([combined_df, filename_df, filtered_df])

# Remove all empty rows from the combined DataFrame
combined_df = combined_df.dropna(how='all')

# Save the combined data to an Excel file in the current directory
combined_file_path = os.path.join(current_directory, output_combined_file)
combined_df.to_excel(combined_file_path, index=False)

print(f"Combined data saved to {combined_file_path}")