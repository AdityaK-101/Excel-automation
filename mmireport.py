# Importing the libraries
import warnings
import pandas as pd
from datetime import datetime, timedelta

df = pd.read_excel("ExportReport.xls", skiprows = 4)
df.head()

status_counts = df['Status (Ticket)'].value_counts()
status_table = status_counts.reset_index()
status_table.columns = ['Status (Ticket)', 'Total']
status_table.head(10)

clarification_tickets = list(map(int, df[df['Status (Ticket)'] == 'Clarification']['Ticket Id'].to_list()))
clarification_tickets.sort()
# print(clarification_tickets, len(clarification_tickets))

resolved_tickets = list(map(int,df[df['Status (Ticket)'] == 'Resolved']['Ticket Id'].to_list()))
resolved_tickets.sort()
# print(resolved_tickets, len(resolved_tickets))

# Calculate 11am yesterday and 11am today
today_11am = datetime.now().replace(hour=11, minute=0, second=0, microsecond=0)
yesterday_11am = today_11am - timedelta(days=2)

# Convert 'created_at' column to datetime if not already
df['Created Time (Ticket)'] = pd.to_datetime(df['Created Time (Ticket)'])

# Filter tickets created between 11am yesterday and 11am today
new_tickets = list(map(int, df[(df['Created Time (Ticket)'] >= yesterday_11am) & (df['Created Time (Ticket)'] < today_11am)]['Ticket Id'].to_list()))
new_tickets.sort()
# print(new_tickets, len(new_tickets))

df.replace('-', None, inplace=True)
# Convert 'created_at' column to datetime if not already
df['Ticket Closed Time'] = pd.to_datetime(df['Ticket Closed Time'])

# Filter tickets created between 11am yesterday and 11am today
closed_tickets = list(map(int,df[(df['Ticket Closed Time'] >= yesterday_11am) & (df['Ticket Closed Time'] < today_11am)]['Ticket Id'].to_list()))
closed_tickets.sort()
# print(closed_tickets, len(closed_tickets))

data1 = {'Status (Ticket)': ['Newly Opened', 'Closed'],
        'Details': [','.join(map(str, new_tickets)), ','.join(map(str, closed_tickets))],
        'Total': [len(new_tickets), len(closed_tickets)]}
new_closed=pd.DataFrame(data1)

data2 = {'Status (Ticket)': ['Resolved', 'Clarification'],
        'Details': [','.join(map(str, resolved_tickets)), ','.join(map(str, clarification_tickets))],
        'Total': [len(resolved_tickets), len(clarification_tickets)]}
res_clar = pd.DataFrame(data2)
#print(ticket_dataframe)

# write to output.html

with open('output_new.html', 'w') as f:
  f.write("<p>Hi Team, </p>")
  f.write("<br>")
  f.write("<p>The following is the status of the tickets on Zoho Desk this morning:</p>")
  f.write(status_table.to_html(index=False,justify='left'))
  f.write("<br>")
  f.write("<p><bold>Zoho Tickets needing MMI team attention: </bold></p>")
  f.write(res_clar.to_html(index=False,justify='left'))
  f.write("<br>")
  f.write("<p>Activity in the last 24 hours before 11am today for open and closed tickets: </p>")
  f.write(new_closed.to_html(index=False,justify='left'))
  f.write("<br>")
  f.write("<p>Have a great day! </p>")
  f.write("<p>Regards,<br>Prakash.</p>")

f.close()
#files.view('/content/sample_data/output.html')