from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from datetime import datetime
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_and_login_robot_order_website()
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    orders_table = download_and_read_orderfile()
    process_orders(orders_table)
    zip_receipts()


def process_orders(orders_table):
    for order in orders_table:
        print("Processing order number: " + order["Order number"])
        close_annoying_modal()
        fill_in_order(order)
        preview_image_filepath = get_robot_preview_image(order["Order number"])
        submit_order()
        receipt_filepath = get_order_receipt(order["Order number"])
        embed_robot_preview_in_pdf(receipt_filepath, preview_image_filepath)
        


def open_and_login_robot_order_website():
    """Open the order page and login"""
    print("Logging in to order website")
    browser.goto("https://robotsparebinindustries.com/")
    LoginPage = browser.page()
    LoginPage.fill("#username", "maria")
    LoginPage.fill("#password", "thoushallnotpass")
    LoginPage.click("button:text('Log in')")    

def close_annoying_modal():
    """Close annoying popup when logging in"""
    print("Closing annoying pop up")
    OrderPage = browser.page()
    OrderPage.click("button:text('OK')")


def download_and_read_orderfile():
    """Downloads, reads the orderfile and returns the data as a table"""
    print("Downloading and reading orderfile")
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders_table = library.read_table_from_csv(path = "orders.csv", columns=["Order number","Head","Body","Legs","Address"])
    return orders_table

def fill_in_order(order):
    """Fills in the order data"""
    print("Filling in order data")
    order_page = browser.page()
    order_page.select_option("#head", str(order["Head"]))
    order_page.click(f'id=id-body-{order["Body"]}')
    order_page.fill('css=input[placeholder="Enter the part number for the legs"]',str(order["Legs"]))
    order_page.fill('css=input[id="address"]',str(order["Address"]))

def get_robot_preview_image(order_number):
    """Gets the robot preview image and saves to output folder"""
    print("Getting robot preview image and saving to output folder")
    order_page = browser.page()
    order_page.click("#preview")
    preview_image_filepath = "output/receipt/robot_preview_"+order_number+".png"
    order_page.locator('id=robot-preview-image').screenshot(path=preview_image_filepath)
    return preview_image_filepath
    

def submit_order():
    """Submits an order and retries if there is an error"""
    print("Submitting order") 
    order_page = browser.page()
    max_attempts = 10
    attempt = 0
    successfully_submitted = False
    while attempt < max_attempts and not successfully_submitted:
        print(f"Attempt {attempt}/{max_attempts} of submitting order")
        try:
            order_page.click('id=order')
            if not order_page.locator("#receipt").count() > 0:
                print("Found error message after submitting")
                raise ValueError("Found error message after submitting")
            else:
                print("Submittal successful")
                successfully_submitted = True
        except:
            attempt +=1
        
    if successfully_submitted:
        pass
    else:
        raise ValueError("Failed to submit order after trying for 10 times")
    
def get_order_receipt(order_number):
    """Saves order receipt into pdf and returns the filepath"""
    print("Saving order receipt into pdf")
    page = browser.page()
    order_receipt = page.locator("#receipt").inner_html()

    pdf = PDF()
    receipt_filepath = "output/receipt/"+order_number+".pdf"
    pdf.html_to_pdf(order_receipt, receipt_filepath)
    print("receipt saved at: "+receipt_filepath)
    page.locator("#order-another").click()
    return receipt_filepath
    
def embed_robot_preview_in_pdf(receipt_filepath, preview_image_filepath):
    """Embeds (appends) the preview image into the receipt pdf"""
    print("embedding robot preview into pdf")
    pdf = PDF()
    print("preview image found ("+preview_image_filepath+"): " + str(os.path.exists(preview_image_filepath)))
    print("receipt found ("+receipt_filepath+"): " + str(os.path.exists(receipt_filepath)))
    pdf.add_files_to_pdf(files = [preview_image_filepath], target_document=receipt_filepath, append = True)
    

def zip_receipts():
    """Zippes the receipt folder into a zipped file"""
    print("zipping receipts")
    current_datetime = datetime.now().strftime("%Y_%m_%d_%H_%M")
    output_zipfile_filepath = "orders_"+current_datetime+".zip"
    receipts_folderpath = "output/receipt"

    lib = Archive()    
    print("Receipt folder found ("+receipts_folderpath+"): " + str(os.path.exists(receipts_folderpath)))
    lib.archive_folder_with_zip(receipts_folderpath, output_zipfile_filepath)
    print("Successfully zipped receipts into: "+output_zipfile_filepath)