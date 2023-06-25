from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType

import os.path
import zipfile
import json
import csv
import re
from datetime import datetime

zip_file = "C:/Adobe Trials/output.json"

if os.path.isfile(zip_file):
    os.remove(zip_file)

input_pdf = "C:/Users/ikerc/Downloads/InvoicesData/InvoicesData/TestDataSet/output1.pdf"

# Initial setup, create credentials instance.
credentials = Credentials.service_account_credentials_builder()\
    .from_file("D:/AdobeFiles/pdfservices-api-credentials.json") \
    .build()

# Create an ExecutionContext using credentials and create a new operation instance.
execution_context = ExecutionContext.create(credentials)
extract_pdf_operation = ExtractPDFOperation.create_new()

# Set operation input from a source file.
source = FileRef.create_from_local_file(input_pdf)
extract_pdf_operation.set_input(source)

# Build ExtractPDF options and set them into the operation
extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
    .with_element_to_extract(ExtractElementType.TEXT) \
    .build()
extract_pdf_operation.set_options(extract_pdf_options)

# Execute the operation.
result: FileRef = extract_pdf_operation.execute(execution_context)

# Save the result to the specified location.
result.save_as(zip_file)

# Extracting JSON data as a dictionary from the zip file provided by the API.
with zipfile.ZipFile(zip_file, 'r') as archive:
    jsonentry = archive.open('structuredData.json')
    jsondata = jsonentry.read()
    data = json.loads(jsondata)  # data is now a dictionary


def get_address_line1(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[6]/Sub[5]":
            return element["Text"]


def get_address_line2(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[6]/Sub[6]":
            return element["Text"]


def get_customer_email(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[6]/Sub[2]":
            return element["Text"]


def get_customer_name(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[6]/Sub":
            return element["Text"]


def get_customer_phone(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[6]/Sub[4]":
            return element["Text"]


def get_invoice_desc(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/Table/TR/TD/P":
            return element["Text"]


def get_due_date(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[9]":
            date_string = element["Text"]
            date_part = date_string.split(": ")[1]
            return date_part


def get_issue_date(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[3]/Sub[3]":
            return element["Text"]


def get_invoice_number(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[3]/Sub":
            return element["Text"].split("# ")[1]


def get_invoice_tax(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/Table[4]/TR[2]/TD[2]/P":
            return element["Text"]


def get_item_name(element):
    pattern = r"//Document/Sect/Table\[3\]/TR(\[\d+\])?/TD/P"
    if re.match(pattern, element["Path"]):
        return element["Text"]


def get_billdetails_quantity(element):
    pattern = r"//Document/Sect/Table\[3\]/TR(\[\d+\])?/TD\[2\]/P"
    if re.match(pattern, element["Path"]):
        return element["Text"]


def get_billdetails_rates(element):
    pattern = r"//Document/Sect/Table\[3\]/TR(\[\d+\])?/TD\[3\]/P"
    if re.match(pattern, element["Path"]):
        return element["Text"]


def get_invoice_desc(data):
    s = ""
    pattern = r"//Document/Sect/Table/TR(\[\d+\])?/TD/P"
    for element in data["elements"]:
        if re.match(pattern, element["Path"]):
            s += element["Text"]
    return s


def get_zipcode(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[2]/Sub[3]":
            return element["Text"]


def get_business_name(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/Title":
            return element["Text"]


def get_business_desc(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[4]":
            return element["Text"]


def get_country(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[2]/Sub[2]":
            return element["Text"]


def extract_street(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[2]/Sub":
            return element["Text"].split(", ")[0]


def extract_city(data):
    for element in data["elements"]:
        if element["Path"] == "//Document/Sect/P[2]/Sub":
            return element["Text"].split(", ")[1]


city = extract_city(data)
street = extract_street(data)
country = get_country(data)
business_desc = get_business_desc(data)
business_name = get_business_name(data)
zipcode = get_zipcode(data)

customer_add_line1 = get_address_line1(data)
customer_add_line2 = get_address_line2(data)
customer_email = get_customer_email(data)
customer_name = get_customer_name(data)
customer_phone = get_customer_phone(data)

invoice_due_date = get_due_date(data)
invoice_issue_date = get_issue_date(data)
invoice_desc = get_invoice_desc(data)
invoice_number = get_invoice_number(data)
invoice_tax = get_invoice_tax(data)


def json_to_csv(json_data, csv_file):
    # Load JSON data
    with open(json_data, 'r') as file:
        data = json.load(file)

    # Extract the desired data from JSON
    names = []
    rows = []
    qtys = []
    rates = []

    for element in data["elements"]:

        if (get_item_name(element) != None):
            names.append(get_item_name(element))
        if (get_billdetails_quantity(element) != None):
            qtys.append(get_billdetails_quantity(element))
        if (get_billdetails_rates(element) != None):
            rates.append(get_billdetails_rates(element))

    i = 0

    for element in data["elements"]:
        if (get_item_name(element) != None):
            row = [city, country, business_desc, business_name, street, zipcode, customer_add_line1, customer_add_line2,
                   customer_email, customer_name, customer_phone, get_item_name(element), qtys[i], rates[i], invoice_desc, invoice_due_date, invoice_issue_date, invoice_number, invoice_tax]
            rows.append(row)
            i += 1


    # Write the data to CSV file
    with open(csv_file, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Bussiness__City', 'Bussiness__Country', 'Bussiness__Description', 'Bussiness__Name', 'Bussiness__StreetAddress', 'Bussiness__Zipcode', 'Customer__Address__line1', 'Customer__Address__line2', 'Customer__Email', 'Customer__Name',
                        'Customer__PhoneNumber', 'Invoice__BillDetails__Name', 'Invoice__BillDetails__Quantity', 'Invoice__BillDetails__Rate', 'Invoice__Description', 'Invoice__DueDate', 'Invoice__IssueDate', 'Invoice__Number', 'Invoice__Tax'])  # Write the header
        writer.writerows(rows)


# Specify the paths for JSON and CSV files
json_file = 'data.json'
csv_file = 'data.csv'

# Write the sample JSON data to a file
with open(json_file, 'w') as file:
    json.dump(data, file)

# Convert JSON to CSV
json_to_csv(json_file, csv_file)
