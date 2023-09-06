# invoice
invoice registration programm 
This is a programm to help people who go shopping every day, and have to register the purchase invoices, we register the invoices in the program through the terminal and they are saved in an Excel document where we can send it to our email once a month if you want or every day, your choice.

mport openpyxl
from openpyxl.styles import Font, Alignment
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Function to create a new invoice entry
def create_invoice():
    date = input("Enter the date (YYYY-MM-DD): ")
    invoice_total = input("Enter the invoice total: ")
    bank_transaction = input("Enter bank transaction details: ")
    commerce_details = input("Enter commerce details: ")

    # Open or create the Excel file
    try:
        workbook = openpyxl.load_workbook("invoices.xlsx")
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Date", "Invoice total", "Bank Transaction", "Commerce Details"])
        for col_num, column_title in enumerate(["A", "B", "C", "D"], 1):
            cell = sheet.cell(row=1, column=col_num)
            cell.value = column_title
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

    # Append the new invoice entry
    sheet.append([date, invoice_total, bank_transaction, commerce_details])

    # Save the workbook
    workbook.save("invoices.xlsx")
    print("Invoice added successfully!")

# Function to send the Excel file via email
def send_email():
    sender_email = "youremail@email.com"
    sender_password = "password"
    receiver_email = "receiveremail@email.cum"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "Monthly Invoices"

    body = "Please find the attached monthly invoices."
    msg.attach(MIMEText(body, 'plain'))

    with open("invoices.xlsx", "rb") as attachment:
        part = MIMEApplication(attachment.read(), _subtype="xlsx")
        part.add_header('content-disposition', 'attachment', filename="invoices.xlsx")
        msg.attach(part)

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Email could not be sent. Error: {str(e)}")

# Main menu
while True:
    print("\nInvoice Management System")
    print("1. Create a new invoice")
    print("2. Send monthly invoices via email")
    print("3. Exit")
    choice = input("Enter your choice (1/2/3): ")

    if choice == "1":
        create_invoice()
    elif choice == "2":
        send_email()
    elif choice == "3":
        print("Exiting the program.")
        break
    else:
        print("Invalid choice. Please try again.")



