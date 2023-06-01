#Script by Will Hoppin for A24 Films, Feb 27th 2023

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
import time
from selenium.webdriver.common.by import By
import os
import re
import shutil
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import numpy as np
import cv2
import pytesseract
import pdf2image
import click
import logging
import datetime
from PIL import Image
from pdf2image import convert_from_path
import PyPDF2


# Create a "PDFs" folder in the same directory as the Python script
pdf_folder = os.path.join(os.path.dirname(__file__), "PDFs")
if not os.path.exists(pdf_folder):
    os.mkdir(pdf_folder)

for file in os.listdir(pdf_folder):
    if file.endswith('.pdf'):
        os.remove(os.path.join(pdf_folder, file))

no_date_folder = os.path.join(os.path.dirname(__file__), "PDFs/NO DATE FOUND")
for file in os.listdir(no_date_folder):
    if file.endswith('.pdf'):
        os.remove(os.path.join(pdf_folder, no_date_folder, file))

#STEP 1: DOWNLOAD PDF FILES

# Set login credentials
username = input("Enter your EP username: ")
password = input("Enter your EP password: ")

date_input = input("Enter the WWE date to download all invoices for in MM/DD/YY format. Example: 03/25/23: ")
print("booting up robochrome - may take a second...")

# Set up Chrome WebDriver
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": os.path.join(os.path.dirname(__file__), "PDFs"),
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True
})

#chrome_options.add_argument('--headless')
service = Service("./chromedriver")
driver = webdriver.Chrome(service=service, options=chrome_options)

# Navigate to login page
driver.get("https://vpo.entertainmentpartners.com/login/default.aspx?ReturnUrl=%2FmainFrameset.aspx")
print("robochrome booted. invoice magic process started. this whole process takes about 25 minutes to run fully... go take your lunch break!")
# Wait for page to load
time.sleep(2)

print("logging in...")

# Fill in login form
username_field = driver.find_element(By.CSS_SELECTOR, 'input[id="LoginControl_UserName"]')
username_field.send_keys(username)

password_field = driver.find_element(By.CSS_SELECTOR, 'input[id="LoginControl_Password"]')
password_field.send_keys(password)

# Submit login form
submit_button = driver.find_element(By.CSS_SELECTOR, 'input[id="LoginControl_Login"]')
submit_button.click()

# Wait for login to complete and display successful login message
driver.implicitly_wait(3)
# switch to the top frame
driver.switch_to.frame('topFrame')
print("navigating to payroll window...")

# locate and click the payroll button
payroll_button = driver.find_element(By.CSS_SELECTOR, 'a[href="/payrolladmin/Default.aspx"] img#payrolladmin')
payroll_button.click()
driver.implicitly_wait(3)

# switch back to the body
driver.switch_to.default_content()
driver.switch_to.frame('body')

print("filtering by given date: " + date_input)
# filter by date
dropdown = driver.find_element(By.ID, "weekEndingDateDropDownList")
dropdown.click()
select = Select(dropdown)
select.select_by_value(date_input)

codes = []
titles = []
#start for loop here
k = 0
while True:
    el_length = len(driver.find_elements(By.CSS_SELECTOR, 'a[id^="PayrollEditDataGrid"]'))
    # save all the titles, and click to expand the invoices!
    for i in range(0,el_length):
        print("opening and downloading invoice package " + str((i+1)+(k*15)) + "...")
        elements = driver.find_elements(By.CSS_SELECTOR, 'a[id^="PayrollEditDataGrid"]')
        titles.append(elements[i].get_attribute("title").strip())
        codes.append(elements[i].text)
        elements[i].click()
        invoice_links = driver.find_elements(By.LINK_TEXT, "Invoice Cover Sheet")
        payroll_links = driver.find_elements(By.LINK_TEXT, "Payroll Register Report")
        fringe_links = driver.find_elements(By.LINK_TEXT, "Fringe Distribution Report")
        residual_due_links = driver.find_elements(By.LINK_TEXT, "Residual Due Audit Report.510")
        for link in invoice_links + payroll_links + fringe_links + residual_due_links:
            link.click()
        elements = driver.find_elements(By.CSS_SELECTOR, 'a[id^="PayrollEditDataGrid"]') 
        elements[i].click()
        time.sleep(5)
        
        print("renaming downloaded files...")
        j = 1
        for filename in os.listdir(pdf_folder):
            if filename.startswith('Display'):
                # extract the code and title from the filename
                code = codes[i+(k*15)]
                title = titles[i+(k*15)]
                # generate the new filename using the code and title
                new_filename = f'{code}_{title}_{j}.pdf'
                # move the file to the new location with the new filename
                shutil.move(os.path.join(pdf_folder, filename), os.path.join(pdf_folder, new_filename))
                j += 1
        
    if el_length != 15:
        break
    else:
        print("moving to next page to repeat process...")
        next_button = driver.find_element(By.XPATH, '//a[contains(text(), "next >>>")]')
        next_button.click()
    k += 1
print("downloading complete! quitting robochrome...")

# Close the browser
driver.quit()

#STEP 2: COMBINE PDF FILES
print("starting to combine pdf files...")
pdf_files = [file for file in os.listdir(pdf_folder) if file.endswith('.pdf')]

pdf_dict = {}
for pdf_file in pdf_files:
    prefix = pdf_file[:6]
    if prefix in pdf_dict:
        pdf_dict[prefix].append(pdf_file)
    else:
        pdf_dict[prefix] = [pdf_file]

# Merge PDF files with the same prefix
for prefix in pdf_dict:
    print("combining files with prefix " + prefix + "...")
    merger = PdfMerger()
    for pdf_file in pdf_dict[prefix]:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        with open(pdf_path, 'rb') as file:
            merger.append(file)

        # Remove PDF file after merging
        os.remove(pdf_path)

    first_pdf_name = pdf_dict[prefix][0]
    output_path = os.path.join(pdf_folder, f"{first_pdf_name[:-6]}.pdf")
    with open(output_path, 'wb') as file:
        merger.write(file)

directory = os.fsencode("PDFs")

print("starting to sort pdf pages based on height to width ratio...")

# Iterate through the combined PDF files
for pdf_file in os.listdir(pdf_folder):
    if pdf_file.endswith('.pdf'):
        pdf_path = os.path.join(pdf_folder, pdf_file)

        # Read PDF file
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)

            # Get the number of pages in the PDF
            num_pages = len(reader.pages)

            # Create a list to store page dimensions
            page_dimensions = []

            # Iterate through pages and get their dimensions
            for i in range(num_pages):
                page = reader.pages[i]
                width = float(page.mediabox.width)
                height = float(page.mediabox.height)
                page_dimensions.append((i, height, width))

            # Calculate height to width ratio and sort pages accordingly
            sorted_pages = sorted(page_dimensions, key=lambda x: x[1] / x[2], reverse=True)

            # Create a new PDF writer
            writer = PdfWriter()

            # Add sorted pages to the new PDF
            for page_info in sorted_pages:
                writer.add_page(reader.pages[page_info[0]])

            # Save the new PDF with sorted pages
            with open(pdf_path, 'wb') as file:
                writer.write(file)

        print(f"sorted pages for {pdf_file}")

print("sorting complete!d.")

def check_for_date(text):
    date = re.search(r"TV SUPPLEMENTAL\s+(\d{2}/\d{2}/\d{4})", text)
    date_matches = re.findall(r"\d{2}/\d{2}/\d{4}", text)
    invoice_date = date_matches[0] if date_matches else None
    if date and invoice_date:
        return date.group(1), invoice_date
    else:
        return None, None

def format_invoice_date(date_string):
    # Convert the date string to a datetime object
    date_obj = datetime.datetime.strptime(date_string, "%m/%d/%Y")
    # Format the date as numbers without separators
    formatted_date = date_obj.strftime("%Y%m%d")
    return formatted_date

folderPath = os.path.join(os.getcwd(), 'PDFs')

# Get a list of PDF files in the specified folder
pdf_files = [file for file in os.listdir(folderPath) if file.endswith('.pdf')]

pdf_files.sort()

new_file_names = []  # List to store the new file names

for pdf_file in pdf_files:
    filePath = os.path.join(folderPath, pdf_file)
    with open(filePath, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        
        if len(pdf_reader.pages) > 0:
            first_page = pdf_reader.pages[0]
            text = first_page.extract_text()
            
            date, invoice_date = check_for_date(text)
            if date and invoice_date:
                old_file_name = os.path.splitext(pdf_file)[0]
                formatted_invoice_date = format_invoice_date(invoice_date)
                new_file_name = f"{old_file_name}_{formatted_invoice_date}_QE_{date.replace('/', '.')}.pdf"
                print("New file name: {}".format(new_file_name))
                new_file_names.append(new_file_name)  # Add new file name to the list
            else:
                print("No date found in '{}'. File not renamed.".format(pdf_file))
        else:
            print("Empty PDF file: '{}'".format(pdf_file))

# Rename the PDF files using the new file names
for old_file_name, new_file_name in zip(pdf_files, new_file_names):
    old_file_path = os.path.join(folderPath, old_file_name)
    new_file_path = os.path.join(folderPath, new_file_name)
    os.rename(old_file_path, new_file_path)
    print("Renamed '{}' to '{}'".format(old_file_name, new_file_name))
