import argparse
import xlrd
import csv
import logging
import time
from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

LOGPATH = ".\mint_loader.log"

REQUIRED_COLS = {"Date", "OriginalDescription", "L1Tag", "L2Tag", "L3Tag", "USD"}
DONOTUPLOAD = 'DNU'

TAGMAP = {
    'Accessories': 'Clothing',
    'Administration - other': 'Fees & Charges',
    'Alcohol': 'Groceries',
    'Appearance': 'Personal Care',
    'Art Supplies': 'Books & Supplies',
    'Bank charges': 'Bank Fee',
    'Beauty products': 'Personal Care',
    'Birthday present': 'Gift',
    'Books & Course Materials': 'Books & Supplies',
    'Books / Magazines / Newspapers': 'Books',
    'Broadband': 'Internet',
    'Business Expenses': 'Business Services',
    'Business Services': 'Business Services',
    'Cash': 'Cash & ATM',
    'Charity - other': 'Charity',
    'Christmas present': 'Gift',
    'Cinema': 'Entertainment',
    'Clothes - Designer or Other': 'Clothing',
    'Clothes - other': 'Clothing',
    'Clothes': 'Clothing',
    'Club Membership': 'Gym',
    'Concert & Theatre': 'Entertainment',
    'Council tax': 'Local Tax',
    'Council Tax': 'Local Tax',
    'Course and Tuition Fees': 'Tuition',
    'Credit Card': 'Fees',
    'Credit card payment': DONOTUPLOAD,
    'Credit card repayment': DONOTUPLOAD,
    'Current account': DONOTUPLOAD,
    'Cycling': 'Public Transportation',
    'Dental treatment': 'Dentist',
    'Designer clothes': 'Clothing',
    'Dining and drinking': 'Restaurants',
    'Dining or Going Out': 'Restaurants',
    'DIY': 'Home Improvement',
    'Domestic supplies': 'Home Supplies',
    'Donation to organisation': 'Charity',
    'Dry cleaning and laundry': 'Laundry',
    'Education - other': 'Education',
    'Electrical equipment': 'Home Supplies',
    'Electricity': 'Utilities',
    'Enjoyment': 'Entertainment',
    'Expenses': 'Reimbursement',
    'Eye care': 'Eyecare',
    'Flights': 'Air Travel',
    'Food, Groceries, Household': 'Groceries',
    'Fuel': 'Gas & Fuel',
    'Furniture': 'Furnishings',
    'Gas and electricity': 'Utilities',
    'Gift': 'Gift',
    'Gifts - other': 'Gift',
    'Gifts or Presents': 'Gift',
    'Groceries': 'Groceries',
    'Gym Equipment': 'Sporting Goods',
    'Gym Membership': 'Gym',
    'Hairdressing': 'Hair',
    'Holiday': 'Travel',
    'Home': 'Home',
    'Home and garden - other': 'Home',
    'Home DIY or Repairs': 'Home Improvement',
    'Home electronics': 'Electronics & Software',
    'Home insurance': 'Home Insurance',
    'Hotel/B&B': 'Hotel',
    'Household - other': 'Home Supplies',
    'Income': 'Income',
    'Insurance - other': 'Financial',
    'Interest charges': 'Finance Charge',
    'Jewellery': 'Shopping',
    'Kitchen / Household Appliances': 'Home',
    'Lighting': 'Furnishings',
    'Lunch or Snacks': 'Restaurants',
    'Media bundle': 'Internet',
    'Medical treatment': 'Doctor',
    'Medication': 'Pharmacy',
    'Mobile': 'Mobile Phone',
    'Museum/exhibition': 'Arts',
    'Office Supplies': 'Office Supplies',
    'Other Repayment': DONOTUPLOAD,
    'Other repayment': DONOTUPLOAD,
    'Parking': 'Parking',
    'Personal Care - Other': 'Personal Care',
    'Personal Electronics': 'Electronics & Software',
    'Pet food': 'Pet Food & Supplies',
    'Pet housing/care': 'Pets',
    'Pets - other': 'Pets',
    'Phone (landline)': 'Home Phone',
    'Physiotherapy': 'Doctor',
    'Postage / Shipping': 'Shipping',
    'Property - other': 'Uncategorized',
    'Public transport': 'Public Transportation',
    'Public Transport': 'Public Transportation',
    'Refunded purchase': 'Returned Purchase',
    'Rent': 'Mortgage & Rent',
    'Rental Income': 'Income',
    'Rental income (room)': 'Income',
    'Rental income (whole property)': 'Income',
    'Repayments': DONOTUPLOAD,
    'Rewards/cashback': 'Uncategorized',
    'Road charges': 'Tolls',
    'Salary (main)': 'Paycheck',
    'Salary (secondary)': 'Paycheck',
    'Salary or Wages (Main)': 'Paycheck',
    'Salary or Wages (Other)': 'Paycheck',
    'Service / Parts / Repairs': 'Service & Parts',
    'Shoes': 'Clothing',
    'Snacks / Refreshments': 'Fast Food',
    'Soft furnishings': 'Furnishings',
    'Software': 'Electronics & Software',
    'Spa': 'Spa & Massage',
    'Sponsorship': 'Charity',
    'Sports Club Membership': 'Gym',
    'Sports Equipment': 'Sporting Goods',
    'Sports event': 'Sports',
    'Stationery & consumables': 'Office Supplies',
    'Supermarket': 'Groceries',
    'Take-away': 'Fast Food',
    'Tax Payment': 'Taxes',
    'Tax rebate': 'Taxes',
    'Taxi': 'Rental Car & Taxi',
    'Toiletries': 'Personal Care',
    'Toys': 'Gift',
    'Transfers': DONOTUPLOAD,
    'Transport - other': 'Travel',
    'Transport': 'Auto & Transport',
    'TV Licence': 'Television',
    'Unsecured Loan repayment': DONOTUPLOAD,
    'Vehicle hire': 'Rental Car & Taxi',
    'Vet': 'Veterinary',
    'Water': 'Utilities',

}

DESC_CUTOFFS = (
    " xx",
    "xxxx",
    "Card:"
)

missing_categories = set()

# code for loading transactions
class Transaction(object):
    def __init__(self, date, desc, category, amount, trans_type="Cash", autodeduct=False, tags=None, notes=None):
        self.date = date
        self.desc = desc
        self.category = category
        self.amount = amount
        self.trans_type = trans_type
        self.autodeduct = autodeduct
        self.tags = tags
        self.notes = notes
        self.income_flag = amount > 0

    def __str__(self):
        return "{dt}\t{desc}\t{cat}\t{amt}".format(dt=self.date,
                                                   desc=self.desc,
                                                   cat=self.category,
                                                   amt=self.amount
                                                   )

    def __lt__(self, other):
        return ((self.date < other.date) or
                (self.date == other.date and self.amount < other.amount)
                )


def _trim_desc(desc):
    for cutoff in DESC_CUTOFFS:
        if cutoff in desc:
            return desc[:desc.find(cutoff)].strip()
    else:
        return desc


def get_tag_translation(tags):
    # first non-null tag will be returned, so they should be passed in priority order
    all_tags = [TAGMAP.get(t, None) for t in tags]

    if not any(all_tags):
        logging.warning("Unmapped category combination detected({}), transaction will not be uploaded".format(tags))
        for t in tags:
            if t:
                missing_categories.add(t)
        raise ValueError()

    return next(t for t in all_tags if t)

def create_transaction_from_moneydashboard_row(row):
    logging.debug("Adding transaction data: {}".format(row))
    try:
        t = Transaction(
            datetime.strptime(row["Date"], "%m/%d/%Y"),
            _trim_desc(row["OriginalDescription"]),
            get_tag_translation([row["L3Tag"], row["L2Tag"], row["L1Tag"]]),
            float(row["USD"].replace(",", ""))
        )
        return t
    except Exception as e:
        logging.exception(e)

def read_transactions(spreadsheet_path, failonmissingcategory):
    if spreadsheet_path.endswith(".xlsx") or spreadsheet_path.endswith(".xls"):
        return read_transactions_from_excel(spreadsheet_path)
    elif spreadsheet_path.endswith(".csv"):
        return read_transactions_from_csv(spreadsheet_path)
    else:
        raise ValueError("Unknown transaction sheet format; please upload in CSV or Excel format")


def read_transactions_from_excel(spreadsheet_path):
    raise NotImplementedError()


def read_transactions_from_csv(spreadsheet_path):
    with open(spreadsheet_path, "r") as csvfile:
        rdr = csv.DictReader(csvfile)
        rows = [r for r in rdr if r["USD"]]
    if not rows:
        raise ValueError("No valid rows found in spreadsheet {}".format(spreadsheet_path))
    elif set(rows[0].keys()).intersection(REQUIRED_COLS) != REQUIRED_COLS:
        raise ValueError("""
        Rows loaded from spreadsheet {} are missing required columns for Transaction creation.  
        Required columns: {}
        Columns detected: {}
        """.format(spreadsheet_path, ", ".join(REQUIRED_COLS), ", ".join(rows[0].keys()))
                         )
    else:
        import pdb; pdb.set_trace()
        transactions = [create_transaction_from_moneydashboard_row(r) for r in rows if r["USD"] != "#N/A"]
        if missing_categories:
            print(missing_categories)
            raise ValueError("Missing transactions detected, halting.")
        # eliminate unmapped 
        transactions = sorted(filter(lambda t: t.category and t.category != DONOTUPLOAD, transactions))
    return transactions


# code for uploading transactions to mint.com   
def add_all_transactions(username, password, transactions):
    def _add_transaction(txn):
        # TODO: check to see if txn exists
        trans_btn = driver.find_element_by_id("controls-add")
        trans_btn.click()
        WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "txnEdit-submit")))
        date_field = driver.find_element_by_id("txnEdit-date-input")
        desc_field = driver.find_element_by_id("txnEdit-merchant_input")
        cat_field = driver.find_element_by_id("txnEdit-category_input")
        amt_field = driver.find_element_by_id("txnEdit-amount_input")
        expense_radio_button = driver.find_element_by_id("txnEdit-mt-expense")
        income_radio_button = driver.find_element_by_id("txnEdit-mt-income")

        # TODO: handle dates outside current year
        date_field.send_keys(Keys.CONTROL, "a", Keys.DELETE)
        date_field.send_keys(txn.date.strftime("%m/%d/%y").upper())
        desc_field.send_keys(txn.desc)
        cat_field.send_keys(txn.category)
        amt_field.send_keys(txn.amount)
        income_radio_button.click() if txn.income_flag else expense_radio_button.click()

        submit_btn = driver.find_element_by_id("txnEdit-submit")
        submit_btn.click()

    driver = webdriver.Chrome()
    driver.get("https://wwws.mint.com")
    driver.implicitly_wait(15)
    el1 = driver.find_element_by_link_text("Log In")
    el1.click()
    # el2 = driver.find_element_by_class_name("mint-auth-login-form")
    # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, "mint-auth-login-form")))
    # el2 = webdriver.support.ui.WebDriverWait(driver, 10).until(
    # EC.presence_of_element_located((By.CLASS_NAME, "mint-auth-login-form")))
    # el2.click()
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "Email")))
    el3 = driver.find_element_by_name("Email")
    el3.send_keys(username.strip())
    el4 = driver.find_element_by_name("Password")
    el4.send_keys(password.strip())
    el5 = driver.find_element_by_name("SignIn")
    el5.click()
    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "transaction")))
    el6 = driver.find_element_by_id("transaction")
    el6.click()
    time.sleep(15)
    for t in transactions:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "controls-add")))
        _add_transaction(t)
        time.sleep(6)

    driver.quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-s", "--spreadsheet_path", default="transactions.xlsx")
    parser.add_argument("-u", "--username")
    parser.add_argument("-p", "--password")
    parser.add_argument("-f", "--failonmissingcategory", default=False, action="store_true")
    parser.add_argument("--load_sheet_only", action="store_true", default=False)
    parser.add_argument("--debug", action="store_true", default=False)
    args = parser.parse_args()

    if not ((args.username and args.password) or args.load_sheet_only):
        raise argparse.ArgumentError(
            "Must pass either username and password or load_sheet_only to prevent uploading to mint.com")

    fmtstring = '%(asctime)s %(levelname)s: %(message)s'
    loglevel = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(filename=LOGPATH, format=fmtstring, level=loglevel)
    console = logging.StreamHandler()
    console.setLevel(loglevel)
    formatter = logging.Formatter(fmtstring)
    console.setFormatter(formatter)
    logging.getLogger("").addHandler(console)

    transactions = read_transactions(args.spreadsheet_path, args.failonmissingcategory)
    if not args.load_sheet_only:
        add_all_transactions(args.username, args.password, transactions)
