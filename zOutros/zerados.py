#C:\Users\Thomas\Downloads\Aleatorio\Zerados



import pandas as pd
import os

# Paths
input_file_path = r'C:\Users\Thomas\Downloads\Base_v2\zOutros\producao.xlsx'
output_folder_path = r'C:\Users\Thomas\Downloads\Base_v2\zOutros\zerados'

# Create the output folder if it doesn't exist
os.makedirs(output_folder_path, exist_ok=True)

# Read the Excel file
df = pd.read_excel(input_file_path)

# Filter rows where the "Total" column equals 0
df_filtered = df[df['Total'] == 0]

# Get a list of unique values in the "Inspetor de producao" column
inspetores_unicos = df_filtered['Inspetor de producao'].unique()

# Loop through each unique "Inspetor de producao" and save filtered data to a new file
for inspetor in inspetores_unicos:
    # Filter rows for the current "Inspetor de producao"
    df_inspetor = df_filtered[df_filtered['Inspetor de producao'] == inspetor][['Inspetor de producao','Corretor']]
    
    # Create a filename for the output Excel file based on the "Inspetor de producao" value
    output_file_path = os.path.join(output_folder_path, f"{inspetor}_zerados.xlsx")
    
    # Save the filtered DataFrame to an Excel file
    df_inspetor.to_excel(output_file_path, index=False)
    print(f"File saved for '{inspetor}' at {output_file_path}")

print("All files have been successfully saved.")