from time import sleep
from robocorp import browser
from robocorp.tasks import task
from datetime import datetime

from RPA.Excel.Files import Files as Excel
from RPA.HTTP import HTTP

# Change value NO to something else to complete the challenge using Javascript
USE_JAVASCRIPT = 'NO'

@task
def solve_challenge():
    HTTP().download("https://aai-devportal-media.s3.us-west-2.amazonaws.com/challenges/humanitarian-response-challenge.xlsx")
    excel = Excel()
    excel.open_workbook("humanitarian-response-challenge.xlsx")
    rows = excel.read_worksheet_as_table("Sheet1", header=True) 

    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )   
    page = browser.goto("https://developer.automationanywhere.com/challenges/humanitarian-response-challenge.html")
    
    if USE_JAVASCRIPT == 'NO':
        for row in rows:
            fill_and_submit_form(row)
        page.click('//button[@id="submit_button"]')
    else:
        for row in rows:
            fill_and_submit_form_js(row)
        page.evaluate(f'''() => {{
            document.evaluate('//button[@id="submit_button"]',document.body,null,9,null).singleNodeValue.click();
        }}''')

    guid_value = page.get_attribute('input[title="Unique GUID"]', 'value')
    print(guid_value)
    browser.screenshot()
    sleep(5)    

def fill_and_submit_form(row):
    page = browser.page()
    size=str(row['Clothing Size']).lower()
    date_string = str(row["Date of Birth"])
    date_obj = datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')   
    formatted_date = date_obj.strftime('%m/%d/%Y') 
    page.fill('//input[@id="customerfirstname"]', row['First Name'])
    page.fill('//input[@id="customerlastname"]', row['Last Name'])
    page.fill('//input[@id="email"]', row['Email'])
    page.fill('//input[@id="city"]', row['City'])
    page.fill('//input[@id="dateofbirth"]', str(formatted_date))
    page.click(f'//input[@id="clothingsize-{size}"]')
    page.select_option('//select[@id="state"]', str(row["State"]))
    page.click('//button[@id="add_button"]')

def fill_and_submit_form_js(row):
    page = browser.page()
    size=str(row['Clothing Size']).lower()
    date_string = str(row["Date of Birth"])
    date_obj = datetime.strptime(date_string, '%Y-%m-%d %H:%M:%S')   
    formatted_date = date_obj.strftime('%m/%d/%Y') 
    page.evaluate(f'''() => {{
        document.evaluate('//input[@id="customerfirstname"]',document.body,null,9,null).singleNodeValue.value='{row['First Name']}';
        document.evaluate('//input[@id="customerlastname"]',document.body,null,9,null).singleNodeValue.value='{row['Last Name']}';
        document.evaluate('//input[@id="email"]',document.body,null,9,null).singleNodeValue.value='{row['Email']}';
        document.evaluate('//input[@id="city"]',document.body,null,9,null).singleNodeValue.value='{row['City']}';
        document.evaluate('//input[@id="dateofbirth"]',document.body,null,9,null).singleNodeValue.value='{str(formatted_date)}';
        document.evaluate('//input[@id="clothingsize-{size}"]',document.body,null,9,null).singleNodeValue.click(); 
        document.evaluate('//select[@id="state"]',document.body,null,9,null).singleNodeValue.value='{str(row["State"])}';
        document.evaluate('//button[@id="add_button"]',document.body,null,9,null).singleNodeValue.click();
    }}''')