import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# --- CONFIGURATION ---
# In a production environment, these might be arguments passed via command line
BUDGET_FILES = [
    "REFS_BUD_SPNSR_NORMN_ORG_1066893487_COPY.xlsx",
    "REFS_BUD_SPNSR_NORMN_ORG_2072399379_COPY.xlsx"
]
PROJECT_INFO_FILE = "OU_SPNSR_ORG_PROJ_264321565_COPY.xlsx"
OUTPUT_FILE = "Combined_Master_Report.xlsx"

def load_and_merge_data():
    """
    ETL Process: Extracts data from multiple Excel sources, transforms types,
    and loads into a single merged DataFrame.
    """
    print("Loading Budget Data...")
    budget_dfs = []
    for file in BUDGET_FILES:
        try:
            # Header is on row 2 (index 1) based on source format
            df = pd.read_excel(file, header=1) 
            budget_dfs.append(df)
        except FileNotFoundError:
            print(f"File {file} not found. Skipping.")

    if not budget_dfs:
        print("No budget files loaded.")
        return pd.DataFrame()

    main_df = pd.concat(budget_dfs, ignore_index=True)

    print("🔄 Loading Project Logic...")
    proj_df = pd.read_excel(PROJECT_INFO_FILE, header=1)

    # Rename columns to match desired output schema
    proj_df = proj_df.rename(columns={
        'Title': 'Grant Title',
        'Sponsor': 'Sponsor Name'
    })

    # Select specific metadata columns to merge
    cols_to_merge = ['Project', 'Sponsor Name', 'Proj Start Date', 'Proj End Date', 'PI Name', 'Grant Title']
    proj_lookup = proj_df[cols_to_merge].drop_duplicates(subset=['Project'])

    # Data Type Normalization (Critical for merging)
    main_df['Project'] = main_df['Project'].astype(str).str.strip()
    proj_lookup['Project'] = proj_lookup['Project'].astype(str).str.strip()

    # Left Join: Keep all financial transactions, attach project metadata where available
    merged_df = pd.merge(main_df, proj_lookup, on='Project', how='left')
    
    # Fill in the Sponsor column from the lookup data
    merged_df['Sponsor'] = merged_df['Sponsor Name']

    return merged_df

def format_excel_report(df):
    # Define Column Order
    final_columns = [
        "Budget Type", "Project", "Fund", "Org", "Sponsor",
        "Proj Start Date", "Proj End Date", "PI Name", "Grant Title",
        "Function", "Account", "Budget Amt", "Pre-Encumbered Amt",
        "Encumbered Amt", "Expended Amt", "Remaining Amt"
    ]

    # Filter and reorder columns
    result_df = pd.DataFrame()
    for col in final_columns:
        result_df[col] = df[col] if col in df.columns else ""

    result_df = result_df.fillna('')

    print(f"Formatting Excel Report ({len(result_df)} rows)...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Financial Summary"

    # Write data to worksheet
    for r_idx, row in enumerate(dataframe_to_rows(result_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Header Styling
            if r_idx == 1:
                cell.font = Font(bold=True, color="FFFFFF", size=11)
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            # Body Styling
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                # Alternating Row Colors
                if r_idx % 2 == 0:
                    cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
                
                # Money Column Formatting (Cols L through P)
                if c_idx >= 12: 
                    cell.number_format = '#,##0.00'

    # Add Borders
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

    # Auto-fit Column Widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column_letter].width = min(max_length + 2, 40)

    # Freeze Header
    ws.freeze_panes = "A2"
    
    wb.save(OUTPUT_FILE)
    print(f"Report saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    combined_data = load_and_merge_data()
    if not combined_data.empty:
        format_excel_report(combined_data)
