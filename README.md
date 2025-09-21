# InvertorCustomerReports

A Python solution to manage inverter customer refills. This repository provides two separate programs:

## Programs

1. **Screen Report** (`invertor_report_screen.py`)
   - Generates a detailed report of pending refills
   - Categorizes customers as:
     - Old pending (refill date before today)
     - Within a week
     - Within a month
   - Prints the report on the screen

2. **WhatsApp Report** (`invertor_report_whatsapp.py`)
   - Sends the same categorized report directly to WhatsApp using `pywhatkit`
   - Requires WhatsApp Web login

## Features
- Reads customer data from Excel (`invertor_customer_data.xlsx`)
- Automatically filters by status and dates
- Clean and readable report formatting
- Flexible: choose to display on screen or send via WhatsApp

## Setup & Installation
1. Install Python 3.x
2. Install required packages:
```bash
pip install pandas pywhatkit
Screen Report
python invertor_report_screen.py
Whatsapp Report
python invertor_report_whatsapp.py
Note: WhatsApp report requires WhatsApp Web login on your browser.

Sample Output
Old Pending Customer Refilling
-----------------------------
Name: Premanandhan N
Address: 123, MG Road, Chennai
Phone: +919876543210
Invoice: INV1234
Purchase Date: 15-Sep-2025
Refill Date: 15-Mar-2026

Customers to be Refilled Within a Week
--------------------------------------
Name: S. Ramesh
Address: 45, Anna Street, Chennai
Phone: +919812345678
Invoice: INV5678
Purchase Date: 21-Sep-2025
Refill Date: 28-Sep-2025

Customers to be Refilled Within a Month
---------------------------------------
Name: K. Priya
Address: 12, Park Avenue, Chennai
Phone: +919898765432
Invoice: INV9012
Purchase Date: 10-Sep-2025
Refill Date: 10-Oct-2025

