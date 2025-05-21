import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Ticket CSV to Excel Converter")

st.title("🎫 Ticket CSV to Excel Converter")
st.write("Upload a CSV file, and get a processed Excel file in return.")

# File uploader
uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

if uploaded_file is not None:
    # Read the CSV file
    df = pd.read_csv(uploaded_file)

    # Map status
    status_translation = {
        'Open': 'Dalam Proses',
        'Resolved': 'Selesai',
        'Closed': 'Tutup'
    }
    df['Current Status'] = df['Current Status'].map(status_translation)

    # Format dates
    date_format_input = '%Y-%m-%d %H:%M:%S'
    df['Date Created'] = pd.to_datetime(df['Date Created'], format=date_format_input, errors='coerce')
    df['Closed Date'] = pd.to_datetime(df['Closed Date'], format=date_format_input, errors='coerce')

    # Sort by date created
    df = df.sort_values(by='Date Created')

    # Format for display
    df['Date Created'] = df['Date Created'].dt.strftime('%d/%m/%Y %I:%M %p')
    df['Closed Date'] = df['Closed Date'].dt.strftime('%d/%m/%Y %I:%M %p')

    # Resolved date logic
    df['Resolved Date'] = df.apply(
        lambda row: row['Closed Date'] if row['Current Status'] == 'Selesai' else ('N/A' if row['Current Status'] == 'Dalam Proses' else ''),
        axis=1
    )

    # Closed date logic
    df['Closed Date'] = df.apply(
        lambda row: 'N/A' if row['Current Status'] in ['Dalam Proses', 'Selesai'] else row['Closed Date'],
        axis=1
    )

    # Add blank columns
    df['Description'] = ''
    df['Respond'] = ''

    # Add serial number
    df.insert(0, 'Bil.', range(1, len(df) + 1))

    # Final column order
    columns = [
        'Bil.',
        'Ticket Number',
        'From',
        'Date Created',
        'Subject',
        'Description',
        'Agent Assigned',
        'Respond',
        'Current Status',
        'Resolved Date',
        'Closed Date',
        'Last Updated',
        'Help Topic'
    ]
    df = df[columns]

    st.success("CSV file processed successfully!")
    st.dataframe(df)

    # Create Excel file in memory
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        label="📥 Download Excel File",
        data=output,
        file_name="converted_tickets.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
