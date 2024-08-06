import pandas as pd
import os

# Replace 'C:/Users/YourUsername/Desktop/your_file.xlsx' with the actual path to your Excel file
file_path = 'C:/Users/Jiraphan.Detchokul/Documents/project_list.xlsx'

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path, engine='openpyxl')

# Extract the first 7 characters from each cell in the third column
extracted_texts = df.iloc[:, 2].apply(lambda x: str(x)[:7])  # Ensure the value is converted to string
print(extracted_texts)

# Format the extracted texts
formatted_texts = ' OR '.join([f'summary ~ "{text}"' for text in extracted_texts])

# Define the path for the new text file
output_file_path = os.path.join(os.path.dirname(file_path), 'extracted_texts.txt')

# Write the formatted text to the new text file
with open(output_file_path, 'w') as file:
    file.write(formatted_texts)

print(f'Formatted text has been written to {output_file_path}')