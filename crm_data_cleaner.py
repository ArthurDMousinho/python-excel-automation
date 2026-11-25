import pandas as pd
import numpy as np
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 1. Generate raw, unformatted sample dataset ("RAW DATA")
raw_data = {
    'Full Name': ['  john DOE ', 'JANE smith', 'john doe', '  Michael   Scott', np.nan, 'sarah Connor  ', 'JANE SMITH'],
    'Email': ['JOHN.DOE@email.com ', 'jane@company.com', 'john.doe@email.com', 'mscott@dunder.com', np.nan, 'sconnor@sky.net ', 'jane@company.com'],
    'Phone': ['(555) 123-4567', '555.987.6543', '(555) 123-4567', '555 222 3333', np.nan, '555-444-5555', '555.987.6543'],
    'Country': ['usa', 'United States', 'USA', 'us', np.nan, 'Canada', 'UNITED STATES']
}

df_dirty = pd.DataFrame(raw_data)
df_dirty.to_excel('1_BEFORE_dirty_leads.xlsx', index=False)

# 2. Data Cleaning & Standardization Pipeline ("CLEAN DATA")
df_clean = df_dirty.copy()

# Drop completely empty rows
df_clean = df_clean.dropna(how='all')

# Standardize full names (trim whitespace and apply Title Case)
df_clean['Full Name'] = df_clean['Full Name'].str.strip().str.replace(r'\s+', ' ', regex=True).str.title()

# Normalize email addresses (lowercase and stripped)
df_clean['Email'] = df_clean['Email'].str.strip().str.lower()

# Standardize phone numbers to (XXX) XXX-XXXX format
df_clean['Phone'] = df_clean['Phone'].str.replace(r'\D', '', regex=True)
df_clean['Phone'] = df_clean['Phone'].apply(lambda x: f"({x[:3]}) {x[3:6]}-{x[6:]}" if len(x) == 10 else x)

# Map inconsistent country abbreviations to official names
country_map = {'usa': 'United States', 'us': 'United States', 'united states': 'United States', 'canada': 'Canada'}
df_clean['Country'] = df_clean['Country'].str.strip().str.lower().map(country_map)

# Remove duplicate entries based on unique Email ID
df_clean = df_clean.drop_duplicates(subset=['Email'], keep='first')

# 3. Export to Excel with Auto-Fit Columns and Clean Corporate Styling
output_file = '2_AFTER_clean_leads.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_clean.to_excel(writer, index=False, sheet_name='Clean_Leads')
    worksheet = writer.sheets['Clean_Leads']

    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    cell_font = Font(name='Calibri', size=11)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col in worksheet.columns:
        max_length = 0
        col_letter = col[0].column_letter
        
        for cell in col:
            cell.border = thin_border
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.font = cell_font
                cell.alignment = Alignment(horizontal='left', vertical='center')
            
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        worksheet.column_dimensions[col_letter].width = max_length + 5

    worksheet.row_dimensions[1].height = 24
    for row_idx in range(2, len(df_clean) + 2):
        worksheet.row_dimensions[row_idx].height = 20

print("Clean and formatted CRM spreadsheet successfully generated!")
