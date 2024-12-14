from utils.openpyxl_generic import WorkbookProperties, CopyWholeSheet
from utils.openpyxl_format_handler import GenericFormatting

workbook_utils = WorkbookProperties()
generic_formatter = GenericFormatting()

# Step 1: Create File, Workbook, and Worksheet
# Initialize the workbook file path and create the workbook
file_path = workbook_utils.create_file("detailed_report.xlsx")
workbook_utils.create_workbook()

# Create a new worksheet named "Comprehensive_Report"
sheet = workbook_utils.create_worksheet("Comprehensive_Report")

# Optional - Remove Default Sheet
# Remove the default sheet if it exists (e.g., "Sheet") to clean up the workbook
workbook_utils.remove_default_sheet()

# Step 2: Write Data to Worksheet
# Write structured data to the worksheet. This data includes headers and sample rows.
data = [
    ["Sales and Purchase Details"],
    ["Category", "Item", "Quantity", "Price", "Total"],
    ["Sales", "Item A", 5, 20, "=C3*D3"],
    ["Sales", "Item B", 10, 15, "=C4*D4"],
    ["Purchases", "Item C", 7, 30, "=C5*D5"],
    ["Purchases", "Item D", 4, 25, "=C6*D6"],
    ["", "Grand Total", "", "", ""]
]
workbook_utils.write_worksheet(data)

# Step 3: Set Sheet Name
# Rename the active sheet to "Detailed_Analysis"
workbook_utils.set_sheet_name("Detailed_Analysis")

# Step 4: Apply Custom Styles
# Configure row and column ranges for styling
workbook_utils.set_row_col()

# Step 5: Apply Borders to Specific Rows
# Apply medium borders to the header row
header_border_request = generic_formatter.form_style_req(2, 1, 5, is_row=True)
workbook_utils.make_border_format(header_border_request, "top_bottom_medium")

# Apply medium borders with right alignment to the "Total" column
total_column_border_request = generic_formatter.form_style_req(2, 5, 6, is_row=True)
workbook_utils.make_border_format(total_column_border_request, "top_bottom_right_medium")

# Step 6: Merge Cells and Apply Color
# Merge the title row cells (row 1, columns 1-5) for a unified title appearance
workbook_utils.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)

# Apply a light blue background color to the title row
title_color = workbook_utils.set_custom_color('lightBlue')
workbook_utils.add_title_color(title_color)

# Step 7: Add Filters to Data
# Add auto-filters to the header row for easier sorting and filtering
workbook_utils.add_filter()

# Step 8: Fetch Row and Column Numbers
# Identify the row number for "Grand Total" in column 2 and the column number for "Total" in row 2
row_number = workbook_utils.fetch_row_number("Grand Total", col_num=2)
column_number = workbook_utils.fetch_column_number("Total", row_num=2)
print(f"Row for 'Grand Total': {row_number}, Column for 'Item': {column_number}")

# Step 9: Use Font Colors
# Set the font color to red for the cell containing "Grand Total"
workbook_utils.set_font_color(row_number, column_number, "red")

# Fill the cell containing "Grand Total" with a light gray background
workbook_utils.fill_cell_color(row_number, column_number, workbook_utils.set_custom_color("lightGray"))

# Step 10: Apply Formulas
# Add a formula to calculate the sum of the "Total" column for rows 3-6
formula_request = generic_formatter.form_formula_req(
    input_row=7, input_col=5, start_row=3, start_col=5, end_row=6, end_col=5
)
workbook_utils.apply_formula(formula_request, "sum")

# Step 11: Apply Styles to Specific Rows
# Apply a bold center style to header rows
header_style_request = generic_formatter.form_style_req(2, 1, 6, is_row=True)
workbook_utils.set_named_styles()
workbook_utils.make_style_format(header_style_request, "bold_center")

# Apply a bold center style to `Grand Total` value
footer_style_request = generic_formatter.form_style_req(7, 2, 3, is_row=True)
workbook_utils.make_style_format(footer_style_request, "bold_center")

# Step 12: Save the Original Workbook
# Save the workbook after all modifications
workbook_utils.save_workbook(file_path)

print(f"Comprehensive report generated and saved at {file_path}")

# Step 13: Load the Workbook for Further Modifications
# Reload the saved workbook to perform additional operations
workbook_utils.load_workbook("Detailed_Analysis", file_path)

# Step 14: Copy Entire Sheet with All Styles and Formulas
# Create a new worksheet "Copied_Report" and copy all data, styles, and formulas from the original sheet
new_sheet = workbook_utils.create_worksheet("Copied_Report")
sheet_copier = CopyWholeSheet(sheet, new_sheet)
sheet_copier.copy_sheet()

# Step 15: Copy Formulas Between Sheets
# Define a request to copy formulas from one sheet to another
copy_formula_request = generic_formatter.form_copy_formula_req(
    cp_sheet="Detailed_Analysis",
    cp_sheet_input_row=7,
    cp_sheet_input_col=5,
    cp_sheet_start_row=7,
    cp_sheet_start_col=5,
    cp_sheet_end_row=7,
    cp_sheet_end_col=5,
)
workbook_utils.set_formula_cp_row_col(copy_formula_request)
workbook_utils.copy_formula()

# Step 16: Save the Copied Workbook
# Save the workbook with the copied sheet
workbook_utils.save_workbook(file_path)

print(f"Copied detailed report generated and saved at {file_path}")