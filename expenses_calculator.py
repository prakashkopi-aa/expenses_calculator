#!/usr/bin/env python3
# filepath: /Users/prakashkopi/Documents/Personal Code Projects/expenses_calculator/expenses_calculator.py

import json
import pandas as pd
import matplotlib.pyplot as plt
import calendar
import os
import logging
import base64
import openpyxl

from premailer import transform
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Email, Attachment
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime, timedelta

# Set up logging as well as final variables
logging.basicConfig(filename='/Users/prakashkopi/Documents/Personal Code Projects/expenses_calculator/logs/lastRun.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
CONFIG_FILE_PATH = '/Users/prakashkopi/Documents/Personal Code Projects/expenses_calculator/config.json'

# Function to read config file
def read_config():
    with open(CONFIG_FILE_PATH) as f:
        config= json.load(f)
    return config

CONFIG = read_config() # Use throughout the code 

# Function to calculate monthly spending from df
def calculate_monthly_spending(df):
    monthly_spending= df['Amount Spent'].sum()
    df['Category'] = df['Category'].str.upper()  # Convert category names to uppercase
    df['Sub-Category']= df['Sub-Category'].str.upper() # Convert sub-category names to uppercase
    monthly_subcategory_spending= df.groupby(['Category', 'Sub-Category'])['Amount Spent'].sum()
    monthly_category_spending= df.groupby('Category')['Amount Spent'].sum()

    return monthly_spending, monthly_subcategory_spending, monthly_category_spending

# Function to calculate yearly spending from df
def calculate_yearly_spending(all_months_df):
    # Calculate total spending, spending by category, and spending by subcategory for the entire year
    yearly_spending = all_months_df['Amount Spent'].sum()
    yearly_category_spending = all_months_df.groupby('Category')['Amount Spent'].sum()
    yearly_amount_saved = sum(CONFIG.get('Income', {}).values()) * 12 - yearly_spending
    
    return yearly_spending, yearly_category_spending, yearly_amount_saved

# Function to remove headers before displaying
def remove_headers(category):
        data_frame= pd.DataFrame(category)
        data_frame_values= data_frame.to_string(header= False)
        return data_frame_values

# Function to generate PDF file 
def generate_pdf(category_spending, total_spending, total_amount_saved, is_yearly=False):
    # Get current month and year
    today = datetime.now()
    if is_yearly:
        title=f"Yearly Spending Report {today.year}"
        filename= f"yearly_spending_report_{today.year}.pdf"
    else:
        month_year = today.strftime("%B %Y")
        title=f"Monthly Spending Report {month_year}"
        filename= f"monthly_spending_report_{month_year}.pdf"

    directory_path= CONFIG['File Path']['directory_path']

    # Create a new PDF file
    c = canvas.Canvas(f"{directory_path}/{filename}", pagesize=letter)

    # Add title
    c.setFont("Helvetica", 16)
    c.drawString(100, 750, title)

    # Add pie chart for spending by category
    plt.figure(figsize=(6,6))
    category_labels = []
    category_values = []
    for category, value in category_spending.items():
        category_labels.append(category)
        category_values.append(value)
    category_labels = [category.capitalize() for category in category_spending.index]    
    def format_percent_value(pct):
        total = sum(category_values)
        absolute = int(pct/100.*total)
        display_format= "{:.1f}%".format(pct)
        if len(category_labels) < 10:
            if pct > 5:
                display_format= "{:.1f}% (${:.2f})".format(pct, absolute)
        elif pct > 10:
            display_format= "{:.1f}% (${:.2f})".format(pct, absolute)

        return display_format

    plt.pie(category_values, labels=category_labels, autopct=format_percent_value, startangle=140)
    plt.axis('equal')  # Equal aspect ratio ensure pie is drawn as a circle
    plt.title('Yearly Spending by Category' if is_yearly else 'Monthly Spending by Category')

    path= '/Users/prakashkopi/Documents/Personal Code Projects/expenses_calculator/Graphs/PNG/' #path to download .png files
    chart_filename= f"{'yearly' if is_yearly else 'monthly'}_spending_category_pie_chart_{today.year if is_yearly else month_year}.png"
    plt.savefig(f'{path}{chart_filename}')

    # Calculate the coordinates to position the pie chart
    chart_width = 400
    chart_height = 400
    chart_x = (letter[0] - chart_width) / 2  # Center horizontally
    chart_y = 300  # Adjust vertically

    # Draw the pie chart image with adjusted coordinates
    c.drawImage(f"{path}{chart_filename}", chart_x, chart_y, width=chart_width, height=chart_height)

    # Add spacing between the pie chart and other elements
    c.drawString(100, 280, "")  # Empty line for spacing

    # Add original output of the code
    c.setFont("Helvetica", 12)
    c.drawString(100, 250, f"Total {'Yearly' if is_yearly else 'Monthly'} Spending: ${total_spending:.2f}")
    c.drawString(100, 230, f"Total Amount Saved for the {'Year' if is_yearly else 'Month'}: ${total_amount_saved:.2f}")

    # Save pdf file
    c.save()

# Function to generate the email message based on the report type
def generate_email_message(report_type, title):
    if report_type not in ['monthly', 'yearly']:
        raise ValueError("Invalid report type. Must be 'monthly' or 'yearly'.")

    subject = f"Your {title} {report_type.capitalize()} Spending Report"
    body = (
        f"Your {report_type} spending report is attached to this email. "
        "Please view the PDF and .txt files for detailed information.\n\n"
    )
    return subject, body

def load_email_template_and_styles(amount_saved):
    base_path = os.path.dirname(os.path.abspath(__file__))
    html_path = os.path.join(base_path, 'templates', 'email_template.html')
    css_path = os.path.join(base_path, 'templates', 'styles.css')

    with open(html_path, 'r') as file:
        html_content = file.read()
    with open(css_path, 'r') as css_file:
        css_content = css_file.read()


    # Determine the font color
    font_color = get_font_color(amount_saved)
    html_content = html_content.replace("{{font_color}}", font_color)

    # Embed CSS styles into HTML template
    html_content = html_content.replace('<link rel="stylesheet" href="./styles.css">', f'<style>{css_content}</style>')

    # Convert CSS to inline styles
    html_content = transform(html_content)

    return html_content

# Function that determines html font color. 
def get_font_color(amount_saved):
    if amount_saved >= 1500:
        return "green"
    elif 500 <= amount_saved < 1500:
        return "#CCCC00"  # Darker yellow color for better legibility
    elif 0 <= amount_saved < 500:
        return "orange"
    else:
        return "red"

# Function to attach file to email
def attach_file_to_email(email, file_path, file_type):
    if os.path.exists(file_path):
        with open(file_path, 'rb' if file_type == "application/pdf" else 'r') as file:
            file_data = file.read()
            attachment = Attachment(
                file_content=base64.b64encode(file_data if file_type == "application/pdf" else file_data.encode('utf-8')).decode(),
                file_type=file_type,
                file_name=os.path.basename(file_path),
                disposition="attachment"
            )
            email.add_attachment(attachment)

def send_email(amount_saved, subject, body, recipient_email, pdf_filename, txt_filename):
    SENDGRID_API_KEY = CONFIG['Send Grid']['email_api_key']
    sg = SendGridAPIClient(SENDGRID_API_KEY)

    # Embed CSS styles into HTML template
    html_content = load_email_template_and_styles(amount_saved)
    html_content = html_content.replace("{{body}}", body)

    # Create the email message
    email = Mail(
        from_email=recipient_email,
        to_emails=recipient_email,
        subject=subject,
        # plain_text_content=body,
        html_content=html_content
    )

    # Attach PDF and .txt files
    attach_file_to_email(email, pdf_filename, "application/pdf")
    attach_file_to_email(email, txt_filename, "text/plain")

    # Send the email
    try:
        response = sg.send(email)
        if 'Monthly' in subject:
            logging.info(f"Monthly email sent! Status Code: {response.status_code}")
            print(f"Monthly email sent! Status Code: {response.status_code}")
        elif 'Yearly' in subject:
            logging.info(f"Yearly email sent! Status Code: {response.status_code}")
            print(f"Yearly email sent! Status Code: {response.status_code}")
    except Exception as e:
        print(f"Error sending email: {e}")
        logging.error(f"Error sending email: {e}")

# Function that creates excel sheet for next month
def create_new_sheet(new_sheet_name):
    # Load the workbook
    file_path= CONFIG['File Path']['file_path']
    wb = openpyxl.load_workbook(file_path)

    # Check if the new sheet already exists
    if new_sheet_name not in wb.sheetnames:
        # Copy the 'Template' sheet
        template_sheet = wb['Template']
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = new_sheet_name
        wb.save(file_path)
        print(f"New sheet '{new_sheet_name}' created.")
        logging.info(f"New sheet '{new_sheet_name}' created.")
    else:
        print(f"Sheet '{new_sheet_name}' already exists.")
        logging.info(f"Sheet '{new_sheet_name}' already exists.")

# Function that saves content to a file
def save_to_file(directory_path, filename, content):
    with open(f'{directory_path}/{filename}', 'w') as txt_file:
        txt_file.write(content)

# (MAIN METHOD) Function to read data from excel file and calculate spending
def calculate_expenses():
    logging.info("START OF RUN")

    # Get current month and year
    today= datetime.now()
    month_year_title= today.strftime("%B %Y")
    year= today.year
    logging.info(f"Date: {today}, Month/Year: {month_year_title}")
    print('Date:', today)
    print('Month/Year: ', month_year_title)

    # Construct sheet name based on month and year title
    sheet_name= month_year_title

    # File path of excel file
    file_path= CONFIG['File Path']['file_path']
    #file_path= '/Users/prakashkopi/VS Code/monthly_expenses_calculator/monthly_expenses.xlsx' # if you want to hardcode file path

    try:

        # Read data from excel file
        df= pd.read_excel(file_path, sheet_name=sheet_name)

        print('Current sheet: ', sheet_name)
        logging.info(f"Current sheet: {sheet_name}")

        # Calculate total spending, spending by category, and spending by subcategory for the month
        monthly_spending, monthly_subcategory_spending, monthly_category_spending= calculate_monthly_spending(df)

        # Calculate monthly income dynamically from config
        monthly_income= sum(CONFIG.get('Income', {}).values())
        print("Monthly income: ", monthly_income)
        logging.info(f"Monthly income: {monthly_income}")

        # Total amount saved
        monthly_amount_saved= monthly_income - monthly_spending
        logging.info(f"Monthly amount spent: {monthly_spending}")
        logging.info(f"Monthly amount saved: {monthly_amount_saved}")

        # Remove headers before displaying
        monthly_subcategory_spending_values= remove_headers(monthly_subcategory_spending)
        monthly_category_spending_values= remove_headers(monthly_category_spending)

        # Create yearly report
        all_months_df = pd.DataFrame()
        for month in range(1, 13):
            sheet_name = f"{calendar.month_name[month]} {today.year}"

            # Check if sheet exists in excel file
            xls= pd.ExcelFile(file_path)
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                df['Category'] = df['Category'].str.upper()
                df['Sub-Category']= df['Sub-Category'].str.upper()
                all_months_df = pd.concat([all_months_df, df], ignore_index=True)

        # Calculate total spending, spending by category, and spending by subcategory for the entire year
        yearly_spending, yearly_category_spending, yearly_amount_saved= calculate_yearly_spending(all_months_df)

        # Remove headers before displaying 
        yearly_category_spending_values= remove_headers(yearly_category_spending)

        # Generate PDF reports
        directory_path= CONFIG['File Path']['directory_path']
        generate_pdf(monthly_category_spending, monthly_spending, monthly_amount_saved, is_yearly=False)
        generate_pdf(yearly_category_spending, yearly_spending, yearly_amount_saved, is_yearly=True)

        # Save full monthly and yearly category spending to a .txt file and attach to email
        with open(f'{directory_path}/{month_year_title}.txt', 'w') as txt_file:
            txt_file.write("Monthly Sub-Category Spending: \n")
            txt_file.write(monthly_subcategory_spending_values)
            txt_file.write('\n\n')
            txt_file.write('------------------------------------------------------------------\n\n')
            txt_file.write("Monthly Category Spending: \n")
            txt_file.write(monthly_category_spending_values)

        with open(f'{directory_path}/{year}.txt', 'w') as txt_file:
            txt_file.write(yearly_category_spending_values)

        logging.info(f"PDF and .txt files generated at path: {directory_path}")

        # Generate monthly/yearly messages
        monthly_subject, monthly_body= generate_email_message('monthly', month_year_title)
        yearly_subject, yearly_body= generate_email_message('yearly', year)

        # Check if tomorrow's date is the 1st of the next month
        tomorrow = today + timedelta(days=1)
        if tomorrow.month == 1 and tomorrow.day == 1:
            # Send yearly email using SendGrid
            send_email(yearly_amount_saved, yearly_subject, yearly_body, CONFIG['Email']['recipient_email'], f"{directory_path}/yearly_spending_report_{year}.pdf", f"{directory_path}/{year}.txt")
            create_new_sheet(f"January {year + 1}")
        elif tomorrow.day == 1:
            # Send monthly email using SendGrid
            send_email(monthly_amount_saved, monthly_subject, monthly_body, CONFIG['Email']['recipient_email'], f"{directory_path}/monthly_spending_report_{month_year_title}.pdf", f"{directory_path}/{month_year_title}.txt")
            create_new_sheet(f"{tomorrow.strftime('%B')} {tomorrow.year}")
        else:
            # can still send email here for testing purposes
            # send_email(yearly_subject, yearly_body, CONFIG['Email']['recipient_email'], f"{directory_path}/yearly_spending_report_{year}.pdf", f"{directory_path}/{year}.txt")
            # send_email(monthly_subject, monthly_body, CONFIG['Email']['recipient_email'], f"{directory_path}/monthly_spending_report_{month_year_title}.pdf", f"{directory_path}/{month_year_title}.txt")
            logging.info("No message sent.")
            print("No message sent.")

    except FileNotFoundError:
        print(f"No expenses recorded for {month_year_title}.")

    logging.info("END OF RUN")

calculate_expenses()