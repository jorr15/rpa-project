# Insert your python code here
while True:
    directory = single_input("Please input the directory where `ProductData.xlsx`, `MasterInvoice.pdf`, and `Order` are located.")
    directory = directory.strip('"').strip("'")
    if directory[-1] != '/':
        directory = directory +'/'
    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        break
    else:
        show_message("Error","The folder does have the correct files. Please try again.")


product_descriptions = {}
product_prices = {}

descriptions = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Descriptions')
for index,row in descriptions.iterrows():
    unit_number = row['Unit number']
    description = row['Description']
    product_descriptions[unit_number] = description
    
prices = pd.read_excel(directory + 'ProductData.xlsx',sheet_name='Prices')
for index,row in prices.iterrows():
    unit_number = row['Unit number']
    price = row['Price per unit']
    product_prices[unit_number] = price

display(product_descriptions)
product_prices

def get_order_data(invoice_num, filename, directory):
   
    template = openpyxl.load_workbook(directory + filename)
    
    # Load the sheet
    selected_sheet = template['OrderForm']

   

    order_data = {}
    order_data['InvoiceNum'] = invoice_num
    date = str(selected_sheet['C9'].value)
    date = date.split()[0].split('-')
    date = f'{date[1]}/{date[2]}/{date[0]}'
    order_data['Date'] = date
    name = selected_sheet['C14'].value
    email = str(selected_sheet['C13'].value)
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}" 
    order_data['Price0'] =  product_prices[selected_sheet['B24'].value]
    order_data['Quantity0'] = selected_sheet['C24'].value
    order_data['Amount0_raw'] = product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value
    order_data['Amount0'] = f'${order_data['Amount0_raw']:,.2f}'
    order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]
    order_data['Total_raw'] = order_data['Amount0_raw']
    order_data['Total'] = f'${order_data['Total_raw']:,.2f}'
    if selected_sheet['B25'].value != None:
        order_data['Price1'] = product_prices[selected_sheet['B25'].value]
        order_data['Quantity1'] = selected_sheet['C25'].value
        order_data['Amount1_raw'] = product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value
        order_data['Amount1'] = f'${order_data['Amount1_raw']:,.2f}'
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]
        order_data['Total_raw'] = order_data['Amount0_raw'] + order_data['Amount1_raw']
        order_data['Total'] = f'${order_data['Total_raw']:,.2f}'
    if selected_sheet['B26'].value != None:
        order_data['Price2'] = product_prices[selected_sheet['B26'].value]
        order_data['Quantity2'] = selected_sheet['C26'].value
        order_data['Amount2_raw'] = product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value
        order_data['Amount2'] = f'${order_data['Amount2_raw']:,.2f}'
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]
        order_data['Total_raw'] = order_data['Amount0_raw'] + order_data['Amount1_raw'] + order_data['Amount2_raw']
        order_data['Total'] = f'${order_data['Total_raw']:,.2f}'
    if selected_sheet['B27'].value != None:
        order_data['Price3'] = product_prices[selected_sheet['B27'].value]
        order_data['Quantity3'] = selected_sheet['C27'].value
        order_data['Amount3_raw'] = product_prices[selected_sheet['B27'].value] * selected_sheet['C27'].value
        order_data['Amount3'] = f'${order_data['Amount3_raw']:,.2f}'
        order_data['Description3'] = product_descriptions[selected_sheet['B27'].value]
        order_data['Total_raw'] = order_data['Amount0_raw'] + order_data['Amount1_raw'] + order_data['Amount2_raw'] + order_data['Amount3_raw']
        order_data['Total'] = f'${order_data['Total_raw']:,.2f}'
    if selected_sheet['B28'].value != None:
        order_data['Price4'] = product_prices[selected_sheet['B28'].value]
        order_data['Quantity4'] = selected_sheet['C28'].value
        order_data['Amount4_raw'] = product_prices[selected_sheet['B28'].value] * selected_sheet['C28'].value
        order_data['Amount4'] = f'${order_data['Amount3_raw']:,.2f}'
        order_data['Description4'] = product_descriptions[selected_sheet['B28'].value]
        order_data['Total_raw'] = order_data['Amount0_raw'] + order_data['Amount1_raw'] + order_data['Amount2_raw'] + order_data['Amount3_raw'] + order_data['Amount4_raw']
        order_data['Total'] = f'${order_data['Total_raw']:,.2f}'
    return order_data, name, email, date

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+ 'MasterInvoice.pdf', directory+ f"Invoice {invoice_num}.pdf", order_data)

def send_email(invoice_num, name, email, date, directory):
    gmail_user = "jakeorr92@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

# List of recipients

    recipients = [email]

# Email subect

    subject = f"Your Order from PythonTech Co. - Invoice {invoice_num}"

# Body of the email

    name = f'{name}'

    body = f"""Dear {name},

    Attached is your Invoice dated {date}.

    Thank You for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84602
    Phonw: (801)555-2372
    Email: sales@pythontechco.com"""

# Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file1 = f'{directory}Invoice {invoice_num}.pdf'

    message = add_attachment(path_to_file1, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


filenames = os.listdir(directory)
filenames = sorted(filenames)
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1