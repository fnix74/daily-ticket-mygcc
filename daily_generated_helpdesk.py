import pandas as pd

# Load the CSV file
input_file = "C:/Users/thefu/Downloads/Daily Ticket Tickets - 20250520.csv"
output_file = "C:/Users/thefu/Downloads/ticket-helpdesk-25-05-20.xlsx"

# Read the CSV into a DataFrame
df = pd.read_csv(input_file)

# Map English status to Malay
status_translation = {
    'Open': 'Dalam Proses',
    'Resolved': 'Selesai',
    'Closed': 'Tutup'
}
df['Current Status'] = df['Current Status'].map(status_translation)

# Define input format
date_format_input = '%Y-%m-%d %H:%M:%S'

# Parse dates first
df['Date Created'] = pd.to_datetime(df['Date Created'], format=date_format_input, errors='coerce')
df['Closed Date'] = pd.to_datetime(df['Closed Date'], format=date_format_input, errors='coerce')

# Sort by 'Date Created' (oldest to newest)
df = df.sort_values(by='Date Created')

# Format date columns after sorting
df['Date Created'] = df['Date Created'].dt.strftime('%d/%m/%Y %I:%M %p')
df['Closed Date'] = df['Closed Date'].dt.strftime('%d/%m/%Y %I:%M %p')

# Add the 'Resolved Date' column
df['Resolved Date'] = df.apply(
    lambda row: row['Closed Date'] if row['Current Status'] == 'Selesai' else ('N/A' if row['Current Status'] == 'Dalam Proses' else ''),
    axis=1
)

# Adjust 'Closed Date'
df['Closed Date'] = df.apply(
    lambda row: 'N/A' if row['Current Status'] in ['Dalam Proses', 'Selesai'] else row['Closed Date'],
    axis=1
)

# Add blank columns
df['Description'] = ''
df['Respond'] = ''

# Add incrementing number
df.insert(0, 'Bil.', range(1, len(df) + 1))

# Rearrange columns
columns = [
    'Bil.', 'Ticket Number', 'From', 'Date Created', 'Subject',
    'Description', 'Agent Assigned', 'Respond', 'Current Status',
    'Resolved Date', 'Closed Date', 'Last Updated', 'Help Topic'
]
df = df[columns]

# Save to Excel
df.to_excel(output_file, index=False)

print("Data has been saved and sorted!")
