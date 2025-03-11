import pandas as pd


pd.set_option('display.max_colwidth', 100)

# The Excel file (  Video A .xslx)
file_path = 'Task2\VIDEOAcopy.xlsx'
file_path_save = 'Task2\VIDEOAcopy_result.xlsx'  
df = pd.read_excel(file_path, sheet_name="Organic")

# Data Frame from 'Video A.xlsx'
print("Original Data:")
print(df.head(10))  


# Here missing values in the specified columns are being interpolate
df['Absolute audience retention (%)'] = df['Absolute audience retention (%)'].interpolate(method='polynomial').round(2) # method= 'linear'
df['Compared to other videos (%)'] = df['Compared to other videos (%)'].interpolate(method='polynomial', order=2).round(2)

# Here to Display the updated DataFrame
print("\nUpdated Data:")
print(df.head(10))  

# Updated DataFrame back to Excel
df.to_excel(file_path_save, index=False, sheet_name="Organic_inter_poly_method")
