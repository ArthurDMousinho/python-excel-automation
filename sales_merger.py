import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 1. Simulate fragmented monthly CSV exports
jan_sales = pd.DataFrame({
    'Order_ID': ['ORD-101', 'ORD-102', 'ORD-103'],
    'Product': ['Mechanical Watch', 'Leather Strap', 'Watch Winder'],
    'Revenue_USD': ['$450.00', '$85.50', '$120.00'],
    'Date': ['2026-01-10', '2026-01-15', '2026-01-22']
})

feb_sales = pd.DataFrame({
    'Order_ID': ['ORD-104', 'ORD-102', 'ORD-105'], # Intentional duplicate Order_ID
    'Product': ['Sapphire Glass', 'Leather Strap', 'Chronograph'],
    'Revenue_USD': ['$60.00', '$85.50', '$890.00'],
    'Date': ['2026-02-03', '2026-01-15', '2026-02-18']
})

jan_sales.to_csv('jan_raw.csv', index=False)
feb_sales.to_csv('feb_raw.csv', index=False)

# 2. Merge, Deduplicate, and Cast Data Types
df_all = pd.concat([pd.read_csv('jan_raw.csv'), pd.read_csv('feb_raw.csv')], ignore_index=True)

# Remove duplicate transactions
df_all = df_all.drop_duplicates(subset=['Order_ID'])

# Convert string currency ('$450.00') into numeric float for Excel calculations
df_all['Revenue_USD'] = df_all['Revenue_USD'].str.replace('$', '', regex=False).astype(float)

# 3. Export Master Report with Corporate Styling & Currency Number Format
output_file = 'MASTER_Q1_Sales_Report.xlsx'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_all.to_excel(writer, index=False, sheet_name='Consolidated_Sales')
    worksheet = writer.sheets['Consolidated_Sales']

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
                # Apply native USD accounting format to Revenue column (Column C)
                if col_letter == 'C':
                    cell.number_format = '$#,##0.00'
            
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        worksheet.column_dimensions[col_letter].width = max_length + 6

    worksheet.row_dimensions[1].height = 24
    for row_idx in range(2, len(df_all) + 2):
        worksheet.row_dimensions[row_idx].height = 20

print("Consolidated Q1 Master Sales Report successfully generated!")
