
import os
import openpyxl
import pandas as pd
import smtplib
from email.message import EmailMessage
from analyticskit import *


# QUESTION 1
files = ["ProductData.xlsx", "MasterInvoice.pdf", "Order1.xlsx",
         "Order2.xlsx", "Order3.xlsx", "Order4.xlsx"]

while True:
    directory = single_input("Enter the directory containing your files")
    directory = directory.strip().replace('"', '')

    if not directory.endswith("\\"):
        directory += "\\"

    if all(os.path.isfile(directory + file) for file in files):
        break
    else:
        show_message("Error", "Invalid directory. Please try again.")


# QUESTION 2
description = pd.read_excel(directory + 'ProductData.xlsx',
                            sheet_name='Descriptions')

product_descriptions = {}

for index, row in description.iterrows():
    unit_number = str(row['Unit number']).strip()
    product_descriptions[unit_number] = str(row['Description'])

prices = pd.read_excel(directory + 'ProductData.xlsx',
                       sheet_name='Prices')

product_prices = {}

for index, row in prices.iterrows():
    unit_number = str(row['Unit number']).strip()
    product_prices[unit_number] = float(row['Price per unit'])


# QUESTION 3
def get_order_data(invoice_num, filename, directory):
    workbook = openpyxl.load_workbook(directory + filename, data_only=True)
    sheet = workbook.active

    date = sheet['C9'].value.strftime('%m/%d/%Y')
    email = str(sheet['C13'].value)
    name = str(sheet['C14'].value)
    company = str(sheet['C15'].value)
    address = str(sheet['C16'].value)
    city = str(sheet['C17'].value)
    state = str(sheet['C18'].value)
    zipcode = str(sheet['C19'].value)

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = company
    order_data['CustomerAddress1'] = address
    order_data['CustomerAddress2'] = f'{city}, {state} {zipcode}'

    total = 0
    row_num = 24
    item_num = 0

    while sheet[f'B{row_num}'].value is not None and item_num < 5:
        unit_number = str(sheet[f'B{row_num}'].value).strip()
        quantity = sheet[f'C{row_num}'].value

        price = product_prices[unit_number]
        description = product_descriptions[unit_number]

        amount = price * quantity
        total += amount

        order_data[f'Description{item_num}'] = str(description)
        order_data[f'Price{item_num}'] = f'${price:,.2f}'
        order_data[f'Quantity{item_num}'] = str(quantity)
        order_data[f'Amount{item_num}'] = f'${amount:,.2f}'

        row_num += 1
        item_num += 1

    order_data['Total'] = f'${total:,.2f}'

    workbook.close()
    return order_data, name, email, date


# QUESTION 4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf',
               directory + f'Invoice {invoice_num}.pdf',
               order_data)


# QUESTION 5
gmail_user = 'william.c.johnston02@gmail.com'

gmail_password_path = select_file("Please select the Gmail password file")
gmail_password = open(gmail_password_path).read().strip().replace(" ", "")


def send_email(invoice_num, name, email, date, directory):
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f'''Dear {name},

Thank you for your order from PythonTech Co. on {date}.

Attached is your invoice for the order.

Thank you for your business!

PythonTech Co.
Email: sales@pythontechco.com'''

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    message = add_attachment(directory + f'Invoice {invoice_num}.pdf',
                             message)

    mail_server = smtplib.SMTP('smtp.gmail.com', 587)
    mail_server.starttls()
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()


# QUESTION 6
filenames = os.listdir(directory)
filenames.sort()
invoice_num = 8515

for filename in filenames:
    if filename.startswith('Order') and filename.endswith('.xlsx'):
        order_data, name, email, date = get_order_data(
            invoice_num, filename, directory
        )

        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)

        invoice_num += 1
