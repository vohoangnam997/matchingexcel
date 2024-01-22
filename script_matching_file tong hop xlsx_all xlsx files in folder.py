import os
import pandas as pd

# User input for column names, file path, and folder path
tong_hop_file_path_input = input("Điền đường dẫn file tổng hợp: ")
column_name_tong_hop_input = input(f"Điền vào tên cột mà bạn muốn match '{tong_hop_file_path_input}': ")
column_name_xlsx_input = input("Điền vào tên cột mà bạn muốn match trong tất cả các files xlsx: ")
xlsx_directory_input = input("Location path của Folder chứa tất cả xlsx file: ")

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

# Iterate over all xlsx files in the specified directory
for xlsx_file_name in os.listdir(xlsx_directory_input):
    if xlsx_file_name.endswith(".xlsx"):
        xlsx_file_path = os.path.join(xlsx_directory_input, xlsx_file_name)

        # Load the xlsx file
        try:
            xlsx_df = pd.read_excel(xlsx_file_path)
        except Exception as e:
            print(f"Error loading xlsx file '{xlsx_file_name}': {e}")
            continue

        # Check if the specified column exists in the xlsx file
        if column_name_xlsx_input not in xlsx_df.columns:
            print(f"Error: '{column_name_xlsx_input}' column not found in xlsx file '{xlsx_file_name}'.")
            continue

        # Iterate through the product names in the xlsx file
        for product_name_xlsx in xlsx_df[column_name_xlsx_input]:
            # Check if the product name exists in "Tong hop phan mem may IT.xlsx" file
            if product_name_xlsx in tong_hop_df[column_name_tong_hop_input].values:
                # Extract user name from xlsx file name
                user_name = os.path.splitext(xlsx_file_name)[0]

                # Update the 'User' column in "Tong hop phan mem may IT.xlsx" with the corresponding user name
                for index, row in tong_hop_df[tong_hop_df[column_name_tong_hop_input] == product_name_xlsx].iterrows():
                    if row['User'] == '':
                        # Check if the product has multiple occurrences
                        if product_name_xlsx in user_assignments:
                            # Distribute user names evenly among matching rows
                            assigned_users = user_assignments[product_name_xlsx]
                            while user_name in assigned_users:
                                user_name += '_duplicate'  # Append '_duplicate' if the user name already exists
                            user_assignments[product_name_xlsx].append(user_name)
                        else:
                            user_assignments[product_name_xlsx] = [user_name]
                        tong_hop_df.at[index, 'User'] = user_name
                        break  # Break the loop after filling one row

# Save the updated DataFrame back to the specified "Tong hop phan mem may IT.xlsx" file
try:
    tong_hop_df.to_excel(tong_hop_file_path_input, index=False)
    print("Update successful.")
except Exception as e:
    print(f"Error saving updated DataFrame to '{tong_hop_file_path_input}': {e}")
    exit(1)
