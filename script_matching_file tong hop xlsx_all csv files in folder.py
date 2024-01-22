import os
import pandas as pd

# User input for column names, file path, and folder path
tong_hop_file_path_input = input("Điền tên file tổng hợp (Excel): ")
column_name_tong_hop_input = input(f"Điền vào tên cột mà bạn muốn match '{tong_hop_file_path_input}': ")
column_name_csv_input = input("Điền vào tên cột mà bạn muốn match trong tất cả các files CSV: ")
csv_directory_input = input("Location path của Folder chứa tất cả CSV file: ")

# Load the "Tong hop phan mem may IT.xlsx" file
try:
    tong_hop_df = pd.read_excel(tong_hop_file_path_input)
except Exception as e:
    print(f"Error loading '{tong_hop_file_path_input}' file: {e}")
    exit(1)

# Check if the specified column exists in "Tong hop phan mem may IT.xlsx"
if column_name_tong_hop_input not in tong_hop_df.columns:
    print(f"Error: '{column_name_tong_hop_input}' column not found in '{tong_hop_file_path_input}'.")
    exit(1)

# Create a new column 'User' in "Tong hop phan mem may IT.xlsx"
tong_hop_df['User'] = ''

# Create a dictionary to keep track of user assignments for each product
user_assignments = {}

# Iterate over all CSV files in the specified directory
for csv_file_name in os.listdir(csv_directory_input):
    if csv_file_name.endswith(".csv"):
        csv_file_path = os.path.join(csv_directory_input, csv_file_name)

        # Load the CSV file
        try:
            csv_df = pd.read_csv(csv_file_path)
        except Exception as e:
            print(f"Error loading CSV file '{csv_file_name}': {e}")
            continue

        # Check if the specified column exists in the CSV file
        if column_name_csv_input not in csv_df.columns:
            print(f"Error: '{column_name_csv_input}' column not found in CSV file '{csv_file_name}'.")
            continue

        # Iterate through the product names in the CSV file
        for product_name_csv in csv_df[column_name_csv_input]:
            # Check if the product name exists in "Tong hop phan mem may IT.xlsx" file
            if product_name_csv in tong_hop_df[column_name_tong_hop_input].values:
                # Extract user name from CSV file name
                user_name = os.path.splitext(csv_file_name)[0]

                # Update the 'User' column in "Tong hop phan mem may IT.xlsx" with the corresponding user name
                for index, row in tong_hop_df[tong_hop_df[column_name_tong_hop_input] == product_name_csv].iterrows():
                    if row['User'] == '':
                        # Check if the product has multiple occurrences
                        if product_name_csv in user_assignments:
                            # Distribute user names evenly among matching rows
                            assigned_users = user_assignments[product_name_csv]
                            while user_name in assigned_users:
                                user_name += '_duplicate'  # Append '_duplicate' if the user name already exists
                            user_assignments[product_name_csv].append(user_name)
                        else:
                            user_assignments[product_name_csv] = [user_name]
                        tong_hop_df.at[index, 'User'] = user_name
                        break  # Break the loop after filling one row

# Save the updated DataFrame back to the specified "Tong hop phan mem may IT.xlsx" file
try:
    tong_hop_df.to_excel(tong_hop_file_path_input, index=False)
    print("Update successful.")
except Exception as e:
    print(f"Error saving updated DataFrame to '{tong_hop_file_path_input}': {e}")
    exit(1)
