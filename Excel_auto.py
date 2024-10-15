#------------------------------------------------------------------excel appending

# import pandas as pd

# xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')
# print(xls.sheet_names)

# df = pd.read_excel(xls, sheet_name='Report1726206738837', header=2, skiprows=2)

# def clarification_search():
#     search_value = 'Clarification'
#     column_to_search = 'Status (Ticket)'
#     filtered_df = df[df[column_to_search] == search_value]
#     clarification_df = pd.DataFrame({'Clarification': filtered_df['Ticket Id'].reset_index(drop=True)})
#     return clarification_df

# def closed_search():
#     search_value = 'Closed'
#     column_to_search = 'Status (Ticket)'
#     filtered_df = df[df[column_to_search] == search_value]
#     closed_df = pd.DataFrame({'Closed': filtered_df['Ticket Id'].reset_index(drop=True)})
#     return closed_df

# def inprogress_search():
#     search_value = 'In Progress'
#     column_to_search = 'Status (Ticket)'
#     filtered_df = df[df[column_to_search] == search_value]
#     inprogress_df = pd.DataFrame({'In Progress': filtered_df['Ticket Id'].reset_index(drop=True)})
#     return inprogress_df

# def onhold_search():
#     search_value = 'On Hold'
#     column_to_search = 'Status (Ticket)'
#     filtered_df = df[df[column_to_search] == search_value]
#     onhold_df = pd.DataFrame({'On Hold': filtered_df['Ticket Id'].reset_index(drop=True)})
#     return onhold_df

# def open_search():
#     search_value = 'Open'
#     column_to_search = 'Status (Ticket)'
#     filtered_df = df[df[column_to_search] == search_value]
#     open_df = pd.DataFrame({'Open': filtered_df['Ticket Id'].reset_index(drop=True)})
#     return open_df

# def resolved_search():
#     search_value = 'Resolved'
#     column_to_search = 'Status (Ticket)'
#     filtered_df = df[df[column_to_search] == search_value]
#     resolved_df = pd.DataFrame({'Resolved': filtered_df['Ticket Id'].reset_index(drop=True)})
#     return resolved_df

# clarification_result = clarification_search()
# closed_result = closed_search()
# inprogress_result = inprogress_search()
# onhold_result = onhold_search()
# open_result = open_search()
# resolved_result = resolved_search()

# result_df = pd.concat([clarification_result, closed_result, inprogress_result, onhold_result, open_result, resolved_result], axis=1)

# with pd.ExcelWriter('ExportReport_1726206739103_20240913_new(AutoRecovered).xlsx', engine='xlsxwriter') as writer:
#     result_df.to_excel(writer, sheet_name='Sheet2', index=False)

# print("Filtered data has been saved.")


#------------------------------------------------------------------rtf with pivot table 1


# import pandas as pd

# # Load Excel file and sheet
# xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')

# # Define required sheet
# required_sheet = 'Report1726206738837'
# if required_sheet not in xls.sheet_names:
#     raise ValueError(f"Sheet named '{required_sheet}' not found in the Excel file.")

# # Load the required sheet into a DataFrame
# df = pd.read_excel(xls, sheet_name=required_sheet, header=2, skiprows=2)

# # Define function to filter tickets based on status and column
# def search_tickets(status_value, column_name, label):
#     filtered_df = df[df[column_name] == status_value]
#     return pd.DataFrame({label: filtered_df['Ticket Id'].reset_index(drop=True)})

# # Parameters for ticket status
# search_params = [
#     ('Clarification', 'Status (Ticket)', 'Clarification'),
#     ('Closed', 'Status (Ticket)', 'Closed'),
#     ('In Progress', 'Status (Ticket)', 'In Progress'),
#     ('On Hold', 'Status (Ticket)', 'On Hold'),
#     ('Open', 'Status (Ticket)', 'Open'),
#     ('Resolved', 'Status (Ticket)', 'Resolved')
# ]

# # Collect results for the statuses
# results = [search_tickets(status, column, label) for status, column, label in search_params]

# # Concatenate the results into one DataFrame
# result_df = pd.concat(results, axis=1)

# # Create a pivot table for status counts
# pivot_df = result_df.melt(var_name='Status', value_name='Ticket Id').dropna()
# pivot_table = pivot_df.groupby('Status').size().reset_index(name='Count')

# # Add a "Grand Total" row
# grand_total = pd.DataFrame({'Status': ['Grand Total'], 'Count': [pivot_table['Count'].sum()]})
# pivot_table = pd.concat([pivot_table, grand_total], ignore_index=True)

# # Function to create a neatly formatted RTF table
# def create_rtf_table_from_df(df):
#     rtf_table = r"{\rtf1\ansi\deff0 {\trowd\trgaph108"  # Begin RTF table
#     cell_def = r"\clbrdrt\brdrs\brdrw10\clbrdrl\brdrs\brdrw10\clbrdrb\brdrs\brdrw10\clbrdrr\brdrs\brdrw10\cellx"
    
#     # Setup columns
#     num_cols = len(df.columns)
#     col_width = 100  # Adjusted column width for better alignment
#     col_defs = ''.join([f"{cell_def}{(i + 1) * col_width}" for i in range(num_cols)])

#     # Add header with blue highlight
#     rtf_table += r"{\highlight1" + col_defs + "\n" + r"".join([f"\intbl {col} \cell " for col in df.columns]) + r"\row}"
    
#     # Add rows
#     for i, row in df.iterrows():
#         row_def = col_defs
#         if i == len(df) - 1 or i == 0:  # Apply blue highlight to first and last rows
#             rtf_table += r"{\highlight1" + row_def + "\n" + r"".join([f"\intbl {cell} \cell " for cell in row]) + r"\row}"
#         else:
#             rtf_table += row_def + "\n" + r"".join([f"\intbl {cell} \cell " for cell in row]) + r"\row"

#     rtf_table += "}}"  # Close the RTF table
#     return rtf_table

# # Generate the RTF table from pivot_table
# rtf_content = create_rtf_table_from_df(pivot_table)

# # Text to include before the table
# intro_text = r"Hi Team,\line\line The following is the status of tickets on Zoho Desk this morning:\line\line"

# # Text after the table
# footer_text = r"\line\line Activity post 11am 12th Sep to 11am 13th Sep:\line"

# # Full RTF content
# full_rtf = r"{\rtf1\ansi\deff0" + intro_text + rtf_content + footer_text + "}"

# # Write RTF content to a file
# with open('daily_ticket_report_new.rtf', 'w') as rtf_file:
#     rtf_file.write(full_rtf)

# print("RTF file with the pivot table has been generated and formatted.")

#------------------------------------------------------------------rtf with formatted pivot table

# import pandas as pd

# # Load Excel file and sheet
# xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')

# # Define required sheet
# required_sheet = 'Report1726206738837'
# if required_sheet not in xls.sheet_names:
#     raise ValueError(f"Sheet named '{required_sheet}' not found in the Excel file.")

# # Load the required sheet into a DataFrame
# df = pd.read_excel(xls, sheet_name=required_sheet, header=2, skiprows=2)

# # Define function to filter tickets based on status and column
# def search_tickets(status_value, column_name, label):
#     filtered_df = df[df[column_name] == status_value]
#     return pd.DataFrame({label: filtered_df['Ticket Id'].reset_index(drop=True)})

# # Parameters for ticket status
# search_params = [
#     ('Clarification', 'Status (Ticket)', 'Clarification'),
#     ('Closed', 'Status (Ticket)', 'Closed'),
#     ('In Progress', 'Status (Ticket)', 'In Progress'),
#     ('On Hold', 'Status (Ticket)', 'On Hold'),
#     ('Open', 'Status (Ticket)', 'Open'),
#     ('Resolved', 'Status (Ticket)', 'Resolved')
# ]

# # Collect results for the statuses
# results = [search_tickets(status, column, label) for status, column, label in search_params]

# # Concatenate the results into one DataFrame
# result_df = pd.concat(results, axis=1)

# # Create a pivot table for status counts
# pivot_df = result_df.melt(var_name='Status', value_name='Ticket Id').dropna()
# pivot_table = pivot_df.groupby('Status').size().reset_index(name='Count')

# # Add a "Grand Total" row
# grand_total = pd.DataFrame({'Status': ['Grand Total'], 'Count': [pivot_table['Count'].sum()]})
# pivot_table = pd.concat([pivot_table, grand_total], ignore_index=True)

# # Function to create a neatly formatted RTF table
# def create_rtf_table_from_df(df):
#     rtf_table = r"{\rtf1\ansi\deff0 {\trowd\trgaph108"  # Begin RTF table
#     cell_def = r"\clbrdrt\brdrs\brdrw10\clbrdrl\brdrs\brdrw10\clbrdrb\brdrs\brdrw10\clbrdrd\brdrs\brdrw10\cellx"
    
#     # Setup columns
#     num_cols = len(df.columns)
#     col_width = 100  # Adjusted column width for better alignment
#     col_defs = ''.join([f"{cell_def}{(i + 1) * col_width}" for i in range(num_cols)])

#     # Add header with blue highlight
#     rtf_table += r"{\highlight1" + col_defs + "\n" + r"".join([f"\intbl {col} \cell " for col in df.columns]) + r"\row}"
    
#     # Add rows
#     for i, row in df.iterrows():
#         rtf_table += col_defs + "\n" + r"".join([f"\intbl {cell} \cell " for cell in row]) + r"\row"

#     rtf_table += "}}"  # Close the RTF table
#     return rtf_table

# # Generate the RTF table from pivot_table
# rtf_content = create_rtf_table_from_df(pivot_table)

# # Text to include before the table
# intro_text = r"Hi Team,\line\line The following is the status of tickets on Zoho Desk this morning:\line\line"

# # Text after the table
# footer_text = r"\line\line Activity post 11am 12th Sep to 11am 13th Sep:\line"

# # Full RTF content
# full_rtf = r"{\rtf1\ansi\deff0" + intro_text + rtf_content + footer_text + "}"

# # Write RTF content to a file
# with open('daily_ticket_report_new_new.rtf', 'w') as rtf_file:
#     rtf_file.write(full_rtf)

# print("RTF file with the pivot table has been generated and formatted.")

#------------------------------------------------------------------rtf with pivot table 2

# import pandas as pd

# # Load Excel file and sheets
# xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')
# xls_sheet2 = pd.ExcelFile('ExportReport_1726206739103_20240913_new(AutoRecovered).xlsx')

# # Define required sheet for the first table
# required_sheet = 'Report1726206738837'
# if required_sheet not in xls.sheet_names:
#     raise ValueError(f"Sheet named '{required_sheet}' not found in the Excel file.")

# # Load the required sheet into a DataFrame
# df = pd.read_excel(xls, sheet_name=required_sheet, header=2, skiprows=2)

# # Define function to filter tickets based on status and column
# def search_tickets(status_value, column_name, label):
#     filtered_df = df[df[column_name] == status_value]
#     return pd.DataFrame({label: filtered_df['Ticket Id'].reset_index(drop=True)})

# # Parameters for ticket status
# search_params = [
#     ('Clarification', 'Status (Ticket)', 'Clarification'),
#     ('Closed', 'Status (Ticket)', 'Closed'),
#     ('In Progress', 'Status (Ticket)', 'In Progress'),
#     ('On Hold', 'Status (Ticket)', 'On Hold'),
#     ('Open', 'Status (Ticket)', 'Open'),
#     ('Resolved', 'Status (Ticket)', 'Resolved')
# ]

# # Collect results for the statuses
# results = [search_tickets(status, column, label) for status, column, label in search_params]

# # Concatenate the results into one DataFrame
# result_df = pd.concat(results, axis=1)

# # Create a pivot table for status counts
# pivot_df = result_df.melt(var_name='Status', value_name='Ticket Id').dropna()
# pivot_table = pivot_df.groupby('Status').size().reset_index(name='Count')

# # Add a "Grand Total" row
# grand_total = pd.DataFrame({'Status': ['Grand Total'], 'Count': [pivot_table['Count'].sum()]})
# pivot_table = pd.concat([pivot_table, grand_total], ignore_index=True)

# # Function to create a neatly formatted RTF table
# def create_rtf_table_from_df(df):
#     rtf_table = r"{\rtf1\ansi\deff0 {\trowd\trgaph108"  # Begin RTF table
#     cell_def = r"\clbrdrt\brdrs\brdrw10\clbrdrl\brdrs\brdrw10\clbrdrb\brdrs\brdrw10\clbrdrd\brdrs\brdrw10\cellx"
    
#     # Setup columns
#     num_cols = len(df.columns)
#     col_width = 100  # Adjusted column width for better alignment
#     col_defs = ''.join([f"{cell_def}{(i + 1) * col_width}" for i in range(num_cols)])

#     # Add header with blue highlight
#     rtf_table += r"{\highlight1" + col_defs + "\n" + r"".join([f"\intbl {col} \cell " for col in df.columns]) + r"\row}"
    
#     # Add rows
#     for i, row in df.iterrows():
#         rtf_table += col_defs + "\n" + r"".join([f"\intbl {cell} \cell " for cell in row]) + r"\row"

#     rtf_table += "}}"  # Close the RTF table
#     return rtf_table

# # Generate the RTF table from pivot_table
# rtf_content_pivot = create_rtf_table_from_df(pivot_table)

# # Load data from Sheet2 for the second table
# sheet2_df = pd.read_excel(xls_sheet2, sheet_name='Sheet2')

# # Filter relevant statuses and ticket numbers from Sheet2
# def create_status_details_table_from_sheet2(df):
#     status_order = ['Clarification', 'Closed', 'Resolved', 'Newly Opened']
#     second_table_df = pd.DataFrame(columns=['Status', 'Details', 'Total'])

#     for status in status_order:
#         if status in df.columns:
#             ticket_numbers = df[status].dropna().astype(int).tolist()  # Get ticket numbers
#             details = ', '.join(map(str, ticket_numbers))  # Concatenate ticket numbers
#             total = len(ticket_numbers)  # Count of tickets
#             second_table_df = second_table_df.append({'Status': status, 'Details': details, 'Total': total}, ignore_index=True)
    
#     return second_table_df

# # Generate the second table DataFrame
# second_table_df = create_status_details_table_from_sheet2(sheet2_df)

# # Generate RTF for the second table
# rtf_content_second_table = create_rtf_table_from_df(second_table_df)

# # Text before the first table
# intro_text = r"Hi Team,\line\line The following is the status of tickets on Zoho Desk this morning:\line\line"

# # Text after the first table and before the second table
# footer_text = r"\line\line Activity post 11am 12th Sep to 11am 13th Sep:\line"

# # Combine RTF content
# full_rtf = r"{\rtf1\ansi\deff0" + intro_text + rtf_content_pivot + footer_text + rtf_content_second_table + "}"

# # Write RTF content to a file
# with open('daily_ticket_report_with_post_activity_table.rtf', 'w') as rtf_file:
#     rtf_file.write(full_rtf)

# print("RTF file with both tables has been generated.")

#------------------------------------------------------------------almost done


import pandas as pd
from datetime import datetime, timedelta

# Load Excel files and sheets
xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')
xls_sheet2 = pd.ExcelFile('ExportReport_1726206739103_20240913_new(AutoRecovered).xlsx')

# Load the required sheet into a DataFrame
required_sheet = 'Report1726206738837'
if required_sheet not in xls.sheet_names:
    raise ValueError(f"Sheet named '{required_sheet}' not found in the Excel file.")
df = pd.read_excel(xls, sheet_name=required_sheet, header=2, skiprows=2)

def clarification_search():
    search_value = 'Clarification'
    column_to_search = 'Status (Ticket)'
    filtered_df = df[df[column_to_search] == search_value]
    clarification_df = pd.DataFrame({'Clarification': filtered_df['Ticket Id'].reset_index(drop=True)})
    return clarification_df

def closed_search():
    search_value = 'Closed'
    column_to_search = 'Status (Ticket)'
    filtered_df = df[df[column_to_search] == search_value]
    closed_df = pd.DataFrame({'Closed': filtered_df['Ticket Id'].reset_index(drop=True)})
    return closed_df

def inprogress_search():
    search_value = 'In Progress'
    column_to_search = 'Status (Ticket)'
    filtered_df = df[df[column_to_search] == search_value]
    inprogress_df = pd.DataFrame({'In Progress': filtered_df['Ticket Id'].reset_index(drop=True)})
    return inprogress_df

def onhold_search():
    search_value = 'On Hold'
    column_to_search = 'Status (Ticket)'
    filtered_df = df[df[column_to_search] == search_value]
    onhold_df = pd.DataFrame({'On Hold': filtered_df['Ticket Id'].reset_index(drop=True)})
    return onhold_df

def open_search():
    search_value = 'Open'
    column_to_search = 'Status (Ticket)'
    filtered_df = df[df[column_to_search] == search_value]
    open_df = pd.DataFrame({'Open': filtered_df['Ticket Id'].reset_index(drop=True)})
    return open_df

def resolved_search():
    search_value = 'Resolved'
    column_to_search = 'Status (Ticket)'
    filtered_df = df[df[column_to_search] == search_value]
    resolved_df = pd.DataFrame({'Resolved': filtered_df['Ticket Id'].reset_index(drop=True)})
    return resolved_df

clarification_result = clarification_search()
closed_result = closed_search()
inprogress_result = inprogress_search()
onhold_result = onhold_search()
open_result = open_search()
resolved_result = resolved_search()

result_df = pd.concat([clarification_result, closed_result, inprogress_result, onhold_result, open_result, resolved_result], axis=1)

with pd.ExcelWriter('ExportReport_1726206739103_20240913_new(AutoRecovered).xlsx', engine='xlsxwriter') as writer:
    result_df.to_excel(writer, sheet_name='Sheet2', index=False)

print("Filtered data has been saved.")

# Function to filter tickets based on status
def search_tickets(status_value, column_name, label):
    filtered_df = df[df[column_name] == status_value]
    return pd.DataFrame({label: filtered_df['Ticket Id'].reset_index(drop=True)})

# Parameters for ticket status
search_params = [
    ('Clarification', 'Status (Ticket)', 'Clarification'),
    ('Closed', 'Status (Ticket)', 'Closed'),
    ('In Progress', 'Status (Ticket)', 'In Progress'),
    ('On Hold', 'Status (Ticket)', 'On Hold'),
    ('Open', 'Status (Ticket)', 'Open'),
    ('Resolved', 'Status (Ticket)', 'Resolved')
]

# Collect results for the statuses
results = [search_tickets(status, column, label) for status, column, label in search_params]
result_df = pd.concat(results, axis=1)

# Create a pivot table for status counts
pivot_table = df.groupby('Status (Ticket)').agg({'Ticket Id': 'count'}).reset_index()
pivot_table.columns = ['Status', 'Count']

# Add a "Grand Total" row
grand_total = pd.DataFrame({'Status': ['Grand Total'], 'Count': [pivot_table['Count'].sum()]})
pivot_table = pd.concat([pivot_table, grand_total], ignore_index=True)

# Function to create a neatly formatted RTF table
def create_rtf_table_from_df(df, col_widths=None):
    rtf_table = r"{\trowd\trgaph108"  # Begin RTF table
    cell_def = r"\cellx"

    # If col_widths are not provided, set default equal widths
    if col_widths is None:
        col_widths = [1500] * len(df.columns)

    # Add column definitions
    col_defs = ''.join([f"{cell_def}{sum(col_widths[:i+1])}" for i in range(len(col_widths))])

    # Add header row (ensure the "Status" header is aligned correctly)
    header_row = r"".join([fr"\intbl {col} \cell " for col in df.columns]) + r"\row"
    rtf_table += col_defs + header_row

    # Add data rows
    for _, row in df.iterrows():
        data_row = r"".join([fr"\intbl {cell} \cell " for cell in row]) + r"\row"
        rtf_table += col_defs + data_row

    rtf_table += "}"  # Close the RTF table
    return rtf_table

# Generate the RTF table from pivot_table
rtf_content = create_rtf_table_from_df(pivot_table)

# Text to include before the first table
intro_text = r"Hi Team,\line\line The following is the status of tickets on Zoho Desk this morning:\line\line"

# Generate dynamic date range
current_date = datetime.now()
previous_date = current_date - timedelta(days=1)

# Format the date range in the format: 11am [previous_date] to 11am [current_date]
previous_date_str = previous_date.strftime('%d %b')
current_date_str = current_date.strftime('%d %b')

# Text after the first table, dynamically changing the date
footer_text = f"\\line\\line Activity post 11am {previous_date_str} to 11am {current_date_str}:\line"

# Full RTF content with the first table
full_rtf = r"{\rtf1\ansi\deff0" + intro_text + rtf_content + footer_text

# Load second sheet for additional table
sheet2_df = pd.read_excel(xls_sheet2, sheet_name='Sheet2')

# Create the second table DataFrame
def create_status_details_table(df):
    status_order = ['Newly Opened', 'Closed', 'Resolved', 'Clarification']
    second_table_df = pd.DataFrame(columns=['Status', 'Details', 'Total'])

    for status in status_order:
        if status in df.columns:
            ticket_numbers = df[status].dropna().astype(int).tolist()  # Get ticket numbers
            details = ', '.join(map(str, ticket_numbers))  # Concatenate ticket numbers
            total = len(ticket_numbers)  # Count of tickets
            new_row = pd.DataFrame({'Status': [status], 'Details': [details], 'Total': [total]})
            second_table_df = pd.concat([second_table_df, new_row], ignore_index=True)

    return second_table_df

# Generate the second table DataFrame
second_table_df = create_status_details_table(sheet2_df)

# Generate RTF content for the second table
rtf_second_table = create_rtf_table_from_df(second_table_df)
full_rtf += rtf_second_table + "}"

# Write RTF content to a file
with open('file.rtf', 'w') as rtf_file:
    rtf_file.write(full_rtf)

print("RTF file with both tables and dynamic dates has been generated and formatted.")




#-------------------------------------------------------------------------------rtf lib not workinh


# import pandas as pd
# from PyRTF import Document, Section, Paragraph, Text, Table, Cell

# # Load Excel file and sheets
# xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')
# xls_sheet2 = pd.ExcelFile('ExportReport_1726206739103_20240913_new(AutoRecovered).xlsx')

# # Load the required sheets into DataFrames
# df = pd.read_excel(xls, sheet_name='Report1726206738837', header=2, skiprows=2)
# sheet2_df = pd.read_excel(xls_sheet2, sheet_name='Sheet2')

# # Define function to filter tickets based on status and column
# def search_tickets(status_value, column_name, label):
#     filtered_df = df[df[column_name] == status_value]
#     return pd.DataFrame({label: filtered_df['Ticket Id']})

# # Parameters for ticket status
# search_params = [
#     ('Clarification', 'Status (Ticket)', 'Clarification'),
#     ('Closed', 'Status (Ticket)', 'Closed'),
#     ('In Progress', 'Status (Ticket)', 'In Progress'),
#     ('On Hold', 'Status (Ticket)', 'On Hold'),
#     ('Open', 'Status (Ticket)', 'Open'),
#     ('Resolved', 'Status (Ticket)', 'Resolved')
# ]

# # Collect results for the statuses
# results = [search_tickets(status, column, label) for status, column, label in search_params]
# result_df = pd.concat(results, axis=1)

# # Create a pivot table for status counts
# pivot_df = result_df.melt(var_name='Status', value_name='Ticket Id').dropna()
# pivot_table = pivot_df.groupby('Status').size().reset_index(name='Count')

# # Add a "Grand Total" row
# grand_total = pd.DataFrame({'Status': ['Grand Total'], 'Count': [pivot_table['Count'].sum()]})
# pivot_table = pd.concat([pivot_table, grand_total], ignore_index=True)

# # Create RTF document and section
# rtf_doc = Document()
# rtf_section = Section()

# # Intro text
# intro_text = Paragraph(Text("Hi Team,\n\nThe following is the status of tickets on Zoho Desk this morning:\n\n"))
# rtf_section.append(intro_text)

# # Function to create RTF table from DataFrame
# def create_rtf_table_from_df(df):
#     table = Table()
    
#     # Add header row
#     table.add_row()
#     table[0][0] = Cell(Text('Status'))
#     table[0][1] = Cell(Text('Total'))
    
#     # Add data rows
#     for _, row in df.iterrows():
#         table.add_row()
#         table[-1][0] = Cell(Text(row['Status']))
#         table[-1][1] = Cell(Text(str(row['Count'])))
    
#     return table

# # Add first table for status counts
# status_table = create_rtf_table_from_df(pivot_table)
# rtf_section.append(status_table)

# # Footer text
# footer_text = Paragraph(Text("\n\nActivity post 11am 12th Sep to 11am 13th Sep:\n"))
# rtf_section.append(footer_text)

# # Create second status details table from Sheet2
# def create_status_details_table_from_sheet2(df):
#     status_order = ['Clarification', 'Closed', 'Resolved', 'Newly Opened']
#     second_table_df = pd.DataFrame(columns=['Status', 'Details', 'Total'])

#     for status in status_order:
#         if status in df.columns:
#             ticket_numbers = df[status].dropna().astype(int).tolist()  # Get ticket numbers
#             details = ', '.join(map(str, ticket_numbers))  # Concatenate ticket numbers
#             total = len(ticket_numbers)  # Count of tickets
#             new_row = pd.DataFrame({'Status': [status], 'Details': [details], 'Total': [total]})
#             second_table_df = pd.concat([second_table_df, new_row], ignore_index=True)
    
#     return second_table_df

# # Generate the second table DataFrame
# second_table_df = create_status_details_table_from_sheet2(sheet2_df)

# # Generate RTF content for the second table
# def create_rtf_second_table_from_df(df):
#     table = Table()
    
#     # Add header row
#     table.add_row()
#     table[0][0] = Cell(Text('Status'))
#     table[0][1] = Cell(Text('Details'))
#     table[0][2] = Cell(Text('Total'))

#     # Add data rows
#     for _, row in df.iterrows():
#         table.add_row()
#         table[-1][0] = Cell(Text(row['Status']))
#         table[-1][1] = Cell(Text(row['Details']))
#         table[-1][2] = Cell(Text(str(row['Total'])))
    
#     return table

# # Add the second table
# second_table = create_rtf_second_table_from_df(second_table_df)
# rtf_section.append(second_table)

# # Append section to document
# rtf_doc.Sections.append(rtf_section)

# # Write the RTF content to a file
# with open('file.rtf', 'w') as rtf_file:
#     rtf_file.write(rtf_doc.Render())

# print("RTF file with both tables has been generated.")





# import pandas as pd
# import pypandoc

# # Load Excel file and sheets
# xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')
# xls_sheet2 = pd.ExcelFile('ExportReport_1726206739103_20240913_new(AutoRecovered).xlsx')

# # Load the required sheets into DataFrames
# df = pd.read_excel(xls, sheet_name='Report1726206738837', header=2, skiprows=2)
# sheet2_df = pd.read_excel(xls_sheet2, sheet_name='Sheet2')

# # Function to filter tickets based on status and column
# def search_tickets(status_value, column_name, label):
#     filtered_df = df[df[column_name] == status_value]
#     return pd.DataFrame({label: filtered_df['Ticket Id'].reset_index(drop=True)})

# # Parameters for ticket status
# search_params = [
#     ('Clarification', 'Status (Ticket)', 'Clarification'),
#     ('Closed', 'Status (Ticket)', 'Closed'),
#     ('In Progress', 'Status (Ticket)', 'In Progress'),
#     ('On Hold', 'Status (Ticket)', 'On Hold'),
#     ('Open', 'Status (Ticket)', 'Open'),
#     ('Resolved', 'Status (Ticket)', 'Resolved')
# ]

# # Collect results for the statuses
# results = [search_tickets(status, column, label) for status, column, label in search_params]
# result_df = pd.concat(results, axis=1)

# # Create a pivot table for status counts
# pivot_df = result_df.melt(var_name='Status', value_name='Ticket Id').dropna()
# pivot_table = pivot_df.groupby('Status').size().reset_index(name='Count')

# # Add a "Grand Total" row
# grand_total = pd.DataFrame({'Status': ['Grand Total'], 'Count': [pivot_table['Count'].sum()]})
# pivot_table = pd.concat([pivot_table, grand_total], ignore_index=True)

# # Create the RTF content as Markdown
# md_content = "# Hi Team\n\n"
# md_content += "The following is the status of tickets on Zoho Desk this morning:\n\n"

# # Create the first table in Markdown
# md_content += "| Status | Total |\n"
# md_content += "|--------|-------|\n"
# for _, row in pivot_table.iterrows():
#     md_content += f"| {row['Status']} | {row['Count']} |\n"

# # Footer text
# md_content += "\n\nActivity post 11am 12th Sep to 11am 13th Sep:\n"

# # Create the second status details table from Sheet2
# def create_status_details_table_from_sheet2(df):
#     status_order = ['Clarification', 'Closed', 'Resolved', 'Newly Opened']
#     second_table_df = pd.DataFrame(columns=['Status', 'Details', 'Total'])

#     for status in status_order:
#         if status in df.columns:
#             ticket_numbers = df[status].dropna().astype(int).tolist()  # Get ticket numbers
#             details = ', '.join(map(str, ticket_numbers))  # Concatenate ticket numbers
#             total = len(ticket_numbers)  # Count of tickets
#             new_row = pd.DataFrame({'Status': [status], 'Details': [details], 'Total': [total]})
#             second_table_df = pd.concat([second_table_df, new_row], ignore_index=True)
    
#     return second_table_df

# # Generate the second table DataFrame
# second_table_df = create_status_details_table_from_sheet2(sheet2_df)

# # Create the second table in Markdown
# md_content += "\n| Status | Details | Total |\n"
# md_content += "|--------|---------|-------|\n"
# for _, row in second_table_df.iterrows():
#     md_content += f"| {row['Status']} | {row['Details']} | {row['Total']} |\n"

# # Convert the Markdown to RTF using pypandoc
# rtf_content = pypandoc.convert_text(md_content, 'rtf', format='md')

# # Write the RTF content to a file
# with open('output.rtf', 'w') as rtf_file:
#     rtf_file.write(rtf_content)

# print("RTF file with both tables has been generated.")



# import pandas as pd
# from docx import Document

# # Load Excel file and sheets
# xls = pd.ExcelFile('ExportReport_1726206739103_20240913_new.xlsx')
# df = pd.read_excel(xls, sheet_name='Report1726206738837', header=2, skiprows=2)
# sheet2_df = pd.read_excel(xls, sheet_name='Sheet2')

# # Collect ticket statuses
# statuses = ['Clarification', 'Closed', 'In Progress', 'On Hold', 'Open', 'Resolved']
# results = {status: df[df['Status (Ticket)'] == status]['Ticket Id'].tolist() for status in statuses}

# # Create a DOCX document
# doc = Document()
# doc.add_heading('Ticket Status Report', level=1)
# doc.add_paragraph("Hi Team,\n\nThe following is the status of tickets on Zoho Desk this morning:\n")

# # Add status information
# for status, tickets in results.items():
#     doc.add_paragraph(f"{status}: {', '.join(map(str, tickets))} (Total: {len(tickets)})")

# doc.add_paragraph("\n\nActivity post 11am 12th Sep to 11am 13th Sep:\n")

# # Additional details from Sheet2
# for index, row in sheet2_df.iterrows():
#     doc.add_paragraph(f"{row['Status']}: {row['Details']} (Total: {row['Total']})")

# # Save the document
# doc.save('ticket_report.docx')
# print("DOCX file has been generated.")
