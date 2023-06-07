# -*- coding: utf-8 -*-
"""
Removing all U+200E charcters from Excel files
"""

import os
import pandas as pd

# Get the current script's directory
script_directory = os.path.dirname(os.path.abspath(__file__))

# Iterate over each file in the directory
for filename in os.listdir(script_directory):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # Load the Excel file
        file_path = os.path.join(script_directory, filename)
        df = pd.read_excel(file_path)

        # Iterate over each cell in the DataFrame
        for index, row in df.iterrows():
            for col in df.columns:
                # Check if the cell contains a string
                if isinstance(df.at[index, col], str):
                    # Remove the U+200E character from the string
                    df.at[index, col] = df.at[index, col].replace('\u200E', '')

        # Save the modified DataFrame back to Excel
        modified_file_path = os.path.join(script_directory, f'modified_{filename}')
        df.to_excel(modified_file_path, index=False)
