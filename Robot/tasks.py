from robocorp.tasks import task
from robocorp import browser
from robocorp.browser import Page
from RPA.Tables import Tables
import os
import requests
import resources.locators as locators
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    csv_url = 'https://robotsparebinindustries.com/orders.csv'
    order_robot_url = 'https://robotsparebinindustries.com/#/robot-order'
    filename = 'orders.csv'
    browser.configure(
        slowmo=500,
    )
    download_csv_file(csv_url, filename)
    open_robot_order_website(order_robot_url)
    close_annoying_modal()
    orders = read_csv_into_table(filename)
    for row in orders:
        fill_the_form(row['Head'],row['Body'],row['Legs'],row['Address'],row['Order number'])
    archive_receipts()
    

def open_robot_order_website(order_robot_url):
    """Navigates to the robot order page"""
    browser.goto(url=order_robot_url)


def download_csv_file(csv_url, filename):
    """
    Download a CSV file from a csv_url and save it to a local file.

    :param csv_url: The csv_url of the CSV file to download.
    :param filename: The name of the file to save the CSV content.
    """
    response = requests.get(csv_url)
    response.raise_for_status()  # This will raise an exception for HTTP errors

    with open(filename, 'wb') as file:
        file.write(response.content)


def read_csv_into_table(csv_url):
    """Reads CSV list into table"""
    library = Tables()
    csv_table = library.read_table_from_csv(csv_url)
    return  csv_table


def close_annoying_modal():
    page = browser.page()
    Page.click(self=page,selector=locators.RSPI_ORDER_NEW_ROBOT_MODAL)


def fill_the_form(head, body, legs, address, order):
    dict = {'1': 'Roll-a-thor body', '2': 'Peanut crusher body', '3': 'D.A.V.E body', '4': 'Andy Roid body', '5': 'Spanner mate body', '6': 'Drillbit 2000 body'}
    body_str = dict.get(body)
    page = browser.page()
    page.select_option("#head", str(head))
    page.get_by_label(body_str).check()
    page.fill(locators.RSPI_ORDER_NEW_ROBOT_TEXT_INPUT_LEGS, legs)
    page.fill("#address", address)
    page.click(selector=locators.RSPI_ORDER_NEW_ROBOT_PREVIEW_BUTTON)
    page.click(selector=locators.RSPI_ORDER_NEW_ROBOT_ORDER_BUTTON)
    x = range(5)
    for n in x:
        alert_visible = page.locator(locators.RSPI_ORDER_NEW_ROBOT_ALERT_TEXT).is_visible()
        if alert_visible:
            page.click(selector=locators.RSPI_ORDER_NEW_ROBOT_ORDER_BUTTON)
        else:
            receipt_visible = page.locator(locators.RSPI_ORDER_NEW_ROBOT_RECEIPT).is_visible()
            if receipt_visible:
                break
    PDF_path = store_receipt_as_pdf(order)
    ss_path = screenshot_robot(order)
    embed_screenshot_to_receipt(ss_path, PDF_path)
    page.click(selector=locators.RSPI_ORDER_NEW_ROBOT_ORDER_ANOTHER_BUTTON)
    close_annoying_modal()


def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator(locators.RSPI_ORDER_NEW_ROBOT_RECEIPT).inner_html()
    pdf = PDF()
    PDF_path = "output/receipts/order_number_" + order_number + "_receipt.pdf"
    pdf.html_to_pdf(sales_results_html, PDF_path)
    return PDF_path


def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    #page.click(selector=locators.RSPI_ORDER_NEW_ROBOT_PREVIEW_BUTTON)
    preview_picture = page.locator(locators.RSPI_ORDER_NEW_ROBOT_PREVIEW)
    ss_path = "output/screenshots/order_number_" + order_number + "_screenshot.png"
    preview_picture.screenshot(type='png',path=ss_path)
    return ss_path


def embed_screenshot_to_receipt(ss_path, pdf_path):
    pdf = PDF()
    list_of_files = [
        ss_path+':align=center'
    ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=pdf_path,
        append=True
        )
    

def archive_receipts():
    Archive.archive_folder_with_zip(self=Archive, folder="output/receipts", archive_name="output/receipts/pdf_receipts.zip")