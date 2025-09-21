import pandas as pd
from datetime import datetime, timedelta
import pywhatkit as kit

# Load Excel file
file_path = "invertor_customer_data.xlsx"
df = pd.read_excel(file_path)

# Convert dates to datetime
df["Invertor Purchase Date"] = pd.to_datetime(df["Invertor Purchase Date"], format="%d-%b-%Y", errors="coerce")
df["Refill Date"] = pd.to_datetime(df["Refill Date"], format="%d-%b-%Y", errors="coerce")

# Today's date
today = datetime.today()

# Filter only pending status
pending_df = df[df["Status"].str.lower() == "pending"]

# Old pending -> refill date before today
old_pending = pending_df[pending_df["Refill Date"] < today]

# Within a week -> refill date between today and 7 days
within_week = pending_df[(pending_df["Refill Date"] >= today) & 
                         (pending_df["Refill Date"] <= today + timedelta(days=7))]

# Within a month -> refill date between 8 and 30 days
within_month = pending_df[(pending_df["Refill Date"] > today + timedelta(days=7)) & 
                          (pending_df["Refill Date"] <= today + timedelta(days=30))]

# Function to format report section
def format_section(title, data):
    report = f"\n{title}\n" + "-"*len(title) + "\n"
    if data.empty:
        report += "No records found.\n"
    else:
        for _, row in data.iterrows():
            report += (f"Name: {row['Name']}\n"
                       f"Address: {row['Address']}\n"
                       f"Phone: {row['Contact Number']}\n"
                       f"Invoice: {row['Invertor Invoice Number']}\n"
                       f"Purchase Date: {row['Invertor Purchase Date'].strftime('%d-%b-%Y')}\n"
                       f"Refill Date: {row['Refill Date'].strftime('%d-%b-%Y')}\n\n")
    return report

# Build full report
report = ""
report += format_section("Old Pending Customer Refilling", old_pending)
report += format_section("Customers to be Refilled Within a Week", within_week)
report += format_section("Customers to be Refilled Within a Month", within_month)

# Send WhatsApp message instantly (requires WhatsApp Web login)
kit.sendwhatmsg_instantly("+919884486609", report, wait_time=20, tab_close=True)
