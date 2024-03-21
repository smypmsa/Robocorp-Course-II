from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

import time

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    orders = get_orders()

    retry_attempt = 0
    max_retry = 3
    
    for order in orders:
        while retry_attempt <= max_retry:
            exception = None
            try:
                close_annoying_modal()
                fill_the_form(order)
                show_preview()
                order_robot()
                handle_error()

                pdf_receipt = store_receipt_as_pdf(order["Order number"])
                screenshot = screenshot_robot(order["Order number"])
                embed_screenshot_to_receipt(screenshot, pdf_receipt)

                order_another_robot() 

            except e:
                exception = e
            finally:
                if exception is None:
                    break
                else:
                    retry_attempt += 1

    archive_receipts("output/receipts")


def open_robot_order_website():
    """Open the robot order website."""
    browser.configure(slowmo=100)
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Read the orders table."""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    library = Tables()
    orders = library.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"])

    for row in orders:
        print(row)

    return orders

def close_annoying_modal():
    """Close the annoying modal appearing after the browser is opened."""
    page = browser.page()
    page.click("text=Yep") if page.is_visible("text=Yep") else None

def fill_the_form(row_data):
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.select_option("id=head", row_data['Head'])
    page.check(f"id=id-body-{row_data['Body']}")
    page.fill("xpath=/html/body/div/div/div[1]/div/div[1]/form/div[3]/input", row_data["Legs"])
    page.fill("xpath=/html/body/div/div/div[1]/div/div[1]/form/div[4]/input", row_data["Address"])

def show_preview():
    page = browser.page()
    page.click("text=Preview")

def order_robot():
    print("Ordering ...")
    page = browser.page()
    page.click(selector="xpath=/html/body/div/div/div[1]/div/div[1]/form/button[2]")

def handle_error():
    page = browser.page()
    if not page.is_visible("xpath=/html/body/div/div/div[1]/div/div[1]/div/button"):
        print("Handling the alert ...")
        while not page.is_visible("xpath=/html/body/div/div/div[1]/div/div[1]/div/button"):
            order_robot()

def order_another_robot():
    page = browser.page()
    page.click(selector="xpath=/html/body/div/div/div[1]/div/div[1]/div/button")

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt_html_locator = page.locator("#receipt").inner_html()

    pdf = PDF()
    pdf_filepath = f"output/receipts/{order_number}.pdf"
    pdf.html_to_pdf(receipt_html_locator, pdf_filepath)

    return pdf_filepath

def screenshot_robot(order_number):
    page = browser.page()
    screenshot_filepath = f"output/screenshots/{order_number}.png"
    screenshot_robot_preview = page.locator("#robot-preview-image").screenshot(path=screenshot_filepath)

    return screenshot_filepath

def embed_screenshot_to_receipt(screenshot_filepath, pdf_filepath):
    pdf = PDF()
    pdf.add_files_to_pdf(files=[pdf_filepath, screenshot_filepath], target_document=pdf_filepath)

def archive_receipts(folder_path):
    archive = Archive()
    archive.archive_folder_with_zip(folder_path, "output/receipts.zip")
