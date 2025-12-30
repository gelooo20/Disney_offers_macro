import os
import xlwings as xw
from datetime import datetime, timedelta

# Indicate start of process
print("Start")

# Get Current directory
current_directory = os.getcwd()


##########################################################################################
# Opening of workbooks
##########################################################################################
first_book = xw.Book('reports.csv')
second_book = xw.Book('Disney Creative Scheduling.xlsx')
third_book = xw.Book('Disney_CreativeQA_Macro[1].xlsm')

# Access the sheets from the workbooks
scheduling_doc = second_book.sheets['FY25_Disney_Creative']
reports_sheet = first_book.sheets['reports']
macro_scheduling_doc = third_book.sheets['GOOGLE DOCS HERE']
macro_reports = third_book.sheets['CREATIVE CHECK DAILY RPT HERE']
macro_page = third_book.sheets['GENERATE REPORTS']

# Scheduling Docs Data
campaign_range = scheduling_doc.range('A:A').value
adConcept_range = scheduling_doc.range('B:B').value
creative_doc_range = scheduling_doc.range('E:E').value
start_range = scheduling_doc.range('G:G').value
end_range = scheduling_doc.range('H:H').value

# Reports Data
date_range = reports_sheet.range('A:A').value
creative_range = reports_sheet.range('B:B').value
creative_id_range = reports_sheet.range('C:C').value
totalads_range = reports_sheet.range('D:D').value

##########################################################################################
# Values loop
##########################################################################################

# First sheet paste
campaign_values = [[item] for item in campaign_range if item is not None]
adconcept_values = [[item] for item in adConcept_range if item is not None]
creative_doc_values = [[item] for item in creative_doc_range if item is not None]
start_values = [[item] for item in start_range if item is not None]
end_values = [[item] for item in end_range if item is not None]

# Second sheet paste
date_values = [[item] for item in date_range if item is not None]
creative_values = [[item] for item in creative_range if item is not None]
creative_id_values = [[item] for item in creative_id_range if item is not None]
total_ads_values = [[item] for item in totalads_range if item is not None]


##########################################################################################
# Paste values from different sheets
##########################################################################################

macro_scheduling_doc.range('A1').value = campaign_values
macro_scheduling_doc.range('B1').value = adconcept_values
macro_scheduling_doc.range('E1').value = creative_doc_values
macro_scheduling_doc.range('F1').value = start_values
macro_scheduling_doc.range('G1').value = end_values

macro_reports.range('A1').value = date_values
macro_reports.range('B1').value = creative_values
macro_reports.range('C1').value = creative_id_values
macro_reports.range('D1').value = total_ads_values

# Change Directory to current directory
macro_page.range('B10').value = current_directory
macro_page.range('I10').value = "Pass 1"

##########################################################################################
# First macro
##########################################################################################

print("Start Macro")

try:

    # Macro Run
    macro = third_book.macro('Module1.GenerateDisneyReports')
    macro()
    print("Running first macro")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    try:
        print("Finished Running first macro")
    except:
        pass  
    
    
##########################################################################################
# Replacement of MAUI Codes Start
##########################################################################################


search_items = ['WDWDOM', 'WDWEPCOT', 'WDWFLRES', 'WDWRSTS', 'WDWDHS', 'DSPNGS', 'WDWRES', 'WDWEEC', 'WDWLATAM', 'CONSUMER', '320x50', '0', 'DIQF']
print("Replacing MAUI CODES")
search_col = macro_scheduling_doc.range('C:C')

# Iterate through the column values
for i, cell_value in enumerate(search_col.value):
    if cell_value in search_items:
        # Copy command
        col_e_value = macro_scheduling_doc.range(f'E{i+1}').value
        
        # Split Command
        split_values = col_e_value.split('_')
        if len(split_values) >= 4: 
            new_value = split_values[3] #Number 3 kasi pang apat yung maui code sa array
        else:
            new_value = ''  # Or handle cases where there are fewer than 4 items

        # Update column C with the 4th element
        macro_scheduling_doc.range(f'C{i+1}').value = new_value
        
# File path and File name
now = datetime.now()
formatted_date = now.strftime('%m.%d.%Y')

##########################################################################################
# Second macro
##########################################################################################


print(f'Current Date: {formatted_date}')

macro_page.range('I10').value = formatted_date + "_Creative_QA_Report"

try:

    # Macro Run
    macro = third_book.macro('Module1.GenerateDisneyReports')
    macro()
    print("Running Second macro")

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    try:
        print("Finished Running Second macro")
    except:
        pass  
    
    
##########################################################################################
# Cleaning the report
##########################################################################################
print("Cleaning the report")


def rgb_to_excel_color(r, g, b):
    return (r << 16) + (g << 8) + b

# Define colors for highlighting
highlight_color_less_than_150 = rgb_to_excel_color(255, 0, 0)  # Red for values less than 150
highlight_color_yesterday = rgb_to_excel_color(255, 255, 0)  # Yellow for yesterday's date

# Get the current date
# formatted_date = datetime.now().strftime('%Y%m%d')  # Format date as needed
file_name = f"{formatted_date}_Creative_QA_Report.xlsx"

# Try to open the Excel workbook
try:
    output_report = xw.Book(file_name)
except FileNotFoundError:
    # Create a new workbook if the file doesn't exist
    output_report = xw.Book()
    output_report.save(file_name)  # Save it with the desired name

# Define the sheets
RemoveRotation = output_report.sheets['Remove From Rotation']
ManualChecking = output_report.sheets['Manual Checking Needed']

# Update cell values in sheets
RemoveRotation.range('A1').value = 'There are no creatives (having a minimum of 150 impressions) in rotation past their end dates.'
ManualChecking.range('A1').value = 'The below creatives have multiple end dates associated with the given MAUI code and job number.'

# Define the range to search within (e.g., the used range of the sheet)
used_range = RemoveRotation.range('G3').current_region

# Get yesterday's date in MM/DD/YYYY format
yesterday_date = (datetime.now() - timedelta(days=1)).date()

# Iterate over each cell in the range
for cell in used_range:
    # Skip empty cells
    if cell.value is None:
        continue

    # Check if the cell value is numeric (for highlighting)
    if isinstance(cell.value, (int, float)):
        if cell.value <= 150:
            cell.api.Interior.Color = highlight_color_less_than_150
    # Check if the cell value is a date
    elif isinstance(cell.value, (datetime, str)):
        # If it's a string, try to convert it to a date
        if isinstance(cell.value, str):
            try:
                cell_date = datetime.strptime(cell.value, '%m/%d/%Y').date()
            except ValueError:
                continue  # Skip if the date format is incorrect
        else:
            cell_date = cell.value.date()  # If it's already a datetime object

        # Highlight corresponding cell in column F if it matches yesterday's date
        if cell_date == yesterday_date:
            cell_offset = cell.offset(0, -1)  # Move to the left to column F
            cell_offset.api.Interior.Color = highlight_color_yesterday

# Save the workbook if needed
output_report.save()
# output_report.close()  # Uncomment if you want to close it after saving

print("Report cleaned.")



##########################################################################################
# Finishing up
##########################################################################################
first_book.close()
second_book.close()

# Save and close the target workbook
output_report.save()

#close all workbook
output_report.close()

print("Program has finished running.")
input("Press Enter to exit...")