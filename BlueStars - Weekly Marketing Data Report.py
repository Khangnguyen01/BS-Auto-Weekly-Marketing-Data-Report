# OS, File Handling, and System Utilities
import os
import sys
import io
import ssl
import zipfile
import pickle
import base64
from io import StringIO, BytesIO
from dotenv import load_dotenv

# Date, Time, and Data Processing
import time
import datetime
from datetime import datetime, timedelta
import json
import re
import numpy as np
import pandas as pd
import pyotp

# Networking and Web Requests
import requests
import urllib.request
from fake_useragent import UserAgent

# Web Scraping and Automation
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# Email Handling
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.message import EmailMessage
from email import encoders

# Google API and Authentication
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ==================================================================================================
#                                         DATA DATE CONFIGURATION
# ==================================================================================================
today = datetime.today().date()

#Period Get Data
report_date = int((datetime.now() - timedelta(days=5)).timestamp())

#Period Data Report
start_of_week = today - timedelta(days=(today.weekday() + 1) % 7)
start_of_last_week = start_of_week - timedelta(days=7)
end_of_last_week = start_of_last_week + timedelta(days=6)

start_date_str = start_of_last_week.strftime('%d.%m')
end_date_str = end_of_last_week.strftime('%d.%m')

# ==================================================================================================
#                                         SETUP GMAIL API AUTHENTICATION
# ==================================================================================================
# Define the scope
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Update paths
JSON_FILE_PATH = '/.../gmail_api.json'
TOKEN_FILE_PATH = '/.../token.pickle'

def authenticate_gmail():
    creds = None
    if os.path.exists(TOKEN_FILE_PATH):
        try:
            with open(TOKEN_FILE_PATH, 'rb') as token:
                creds = pickle.load(token)
        except Exception as e:
            print(f"Error loading credentials: {e}")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing credentials: {e}")
                creds = None
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(JSON_FILE_PATH, SCOPES)
                creds = flow.run_local_server(port=0)
            except Exception as e:
                print(f"Error during authentication: {e}")
                return None
        try:
            with open(TOKEN_FILE_PATH, 'wb') as token:
                pickle.dump(creds, token)
        except Exception as e:
            print(f"Error saving credentials: {e}")

    service = build('gmail', 'v1', credentials=creds)
    return service

# ==================================================================================================
#                                         SETUP EMAIL EXTRACTION FUNCTION
# ==================================================================================================
def extract_hyperlinks_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    links = []
    for link in soup.find_all('a', href=True):
        links.append(link['href'])
    return links

def get_filtered_emails(service, max_results=10):
    subjects = {
        "Weekly BlueStars US Sponsored Products Campaign report": "BS_US_SP_link",
        "Weekly BlueStars US Sponsored Brands Campaign report": "BS_US_SB_link",
        "Weekly BlueStars US Sponsored Display Campaign report": "BS_US_SD_link",
        "Weekly BlueStars CA Sponsored Products Campaign report": "BS_CA_SP_link",
        "Weekly BlueStars CA Sponsored Brands Campaign report": "BS_CA_SB_link",
        "Weekly BlueStars CA Sponsored Display Campaign report": "BS_CA_SD_link"
    }

    links = {key: None for key in subjects.values()}

    for subject, link_var in subjects.items():
        query = f'(from:noreply@amazon.com OR from:no-reply@amazon.com) subject:"{subject}" after:{report_date} before:{today}'

        # Fetch the list of messages matching the query
        results = service.users().messages().list(userId='me', q=query, maxResults=max_results).execute()
        messages = results.get('messages', [])

        if not messages:
            print(f'No matching messages found for subject: {subject}')
            continue

        for msg in messages:
            # Get the message details
            message = service.users().messages().get(userId='me', id=msg['id'], format='full').execute()

            # Extract the body of the email
            for part in message['payload']['parts']:
                if part['mimeType'] == 'text/html':
                    html_content = part['body']['data']
                    # Gmail API returns the body in base64 encoded format, so we need to decode it
                    html_content = html_content.replace('-', '+').replace('_', '/')
                    decoded_html = base64.urlsafe_b64decode(html_content).decode('utf-8')

                    # Extract hyperlinks from the HTML content
                    found_links = extract_hyperlinks_from_html(decoded_html)
                    if found_links:
                        links[link_var] = found_links[0]  # Store the first link
                    else:
                        print(f"No valid download link found for subject: {subject}.")

    return links

# ==================================================================================================
#                                         SETUP DOWNLOAD FUNCTION
# ==================================================================================================
def download_file_to_dataframe(link, session, max_retries=3, delay=5):
    """Tải file từ link và chuyển thành DataFrame (hỗ trợ CSV & XLSX)."""
    for attempt in range(max_retries):
        try:
            response = session.get(link, stream=True)
            if response.status_code != 200:
                raise Exception(f"HTTP {response.status_code}: Không thể tải báo cáo!")

            content = response.content
            try:
                df = pd.read_csv(io.StringIO(content.decode('utf-8')), on_bad_lines='skip', delimiter=',', encoding='utf-8')
                print(f"📄 Đọc thành công file CSV!")
                return df
            except Exception as csv_error:
                print(f"⚠️ Lỗi CSV: {csv_error}, thử XLSX...")

            try:
                df = pd.read_excel(io.BytesIO(content), engine='openpyxl')
                print(f"📄 Đọc thành công file XLSX!")
                return df
            except Exception as excel_error:
                print(f"⚠️ Lỗi XLSX: {excel_error}")

            if attempt < max_retries - 1:
                print(f"🔄 Thử lại sau {delay} giây...")
                time.sleep(delay)

        except Exception as e:
            if '503' in str(e):
                print(f"🚨 Lỗi 503: Lần thử {attempt + 1}/{max_retries}. Đang thử lại sau {delay} giây...")
                time.sleep(delay)
            else:
                print(f"❌ Lỗi tải file: {e}")
                return None

    print(f"❌ Vượt quá số lần thử. Không thể tải hoặc xử lý file từ {link}.")
    return None
# ==================================================================================================
#                                         SETUP MAIN FUNCTION
# ==================================================================================================
if __name__ == '__main__':
    service = authenticate_gmail()
    links = get_filtered_emails(service)

# ==================================================================================================
#                                         GMAIL CREDENTIALS AUTHENTICATION
# ==================================================================================================
load_dotenv('/.../credentials.env')

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
TOTP_SECRET = os.getenv("TOTP_SECRET")

# ==================================================================================================
#                                         ACCESS LINKS AND DOWNLOAD FILES
# ==================================================================================================
def web_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--verbose")
    options.add_argument('--no-sandbox')
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument("--window-size=1920, 1200")
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)
    return driver

driver = web_driver()

try:
    wait = WebDriverWait(driver, 15)

    # MỞ TRANG ĐĂNG NHẬP
    first_report_name, first_report_link = next(iter(links.items()))
    driver.get(first_report_link)
    time.sleep(2)

    # NHẬP EMAIL
    email_input = wait.until(EC.presence_of_element_located((By.ID, "ap_email")))
    email_input.send_keys(EMAIL)
    email_input.send_keys(Keys.RETURN)
    time.sleep(2)

    # NHẬP MẬT KHẨU
    password_input = driver.find_element(By.ID, "ap_password")
    password_input.send_keys(PASSWORD)
    password_input.send_keys(Keys.RETURN)
    time.sleep(2)

    # NHẬP OTP
    totp = pyotp.TOTP(TOTP_SECRET)
    otp_code = totp.now()

    otp_input = driver.find_element(By.ID, "auth-mfa-otpcode")
    otp_input.send_keys(otp_code)
    otp_input.send_keys(Keys.RETURN)
    time.sleep(5)

    # LẤY COOKIES ĐỂ SỬ DỤNG VỚI REQUESTS
    cookies = driver.get_cookies()
    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie["name"], cookie["value"])

    # DUYỆT QUA CÁC LINK VÀ TẢI BÁO CÁO
    dataframes = {}

    for report_name, report_link in links.items():
        print(f"\n📥 Đang tải báo cáo {report_name}: {report_link}")

        df = download_file_to_dataframe(report_link, session)

        if df is not None:
            dataframes[report_name] = df
            print(f"✅ Lưu thành công DataFrame: {report_name}")
        else:
            print(f"❌ Không thể đọc báo cáo {report_name}!")

finally:
    driver.quit()

# ==================================================================================================
#                                         SETUP PROCESS DATAFRAME FUNCTION
# ==================================================================================================
def process_dataframe(df, df_name):
    product_id = None

    # Determine brand
    if 'BS' in df_name:
        brand = 'BlueStars'
        product_id = pd.read_html('https://docs.google.com/spreadsheets/d/e/2PACX-1vQKItZA2bNZCY2yD52UEunCjkuq8e9yDuHLQzLqbOvLy13GeJWLWFmujTCRbrBVNA/pubhtml?gid=678803432&single=true', skiprows=1)[0]
    elif 'CNM' in df_name:
        brand = 'Canamax'
        product_id = pd.read_html('https://docs.google.com/spreadsheets/d/e/2PACX-1vT76uZxbyhrqmCKrR5hODwpvEruqmvbqEfSPvC1S2qCWfwdaHHfqbJpT-lUELcgbw/pubhtml?gid=45330461&single=true', skiprows=1)[0]
    else:
        raise ValueError(f"Không thể xác định thương hiệu từ {df_name}")

    # Determine ad_type
    if 'SP' in df_name:
        ad_type = 'SP'
    elif 'SB' in df_name:
        ad_type = 'SB'
    elif 'SD' in df_name:
        ad_type = 'SD'
    else:
        raise ValueError(f"Không thể xác định loại quảng cáo từ {df_name}")

    # Determine market
    if 'US' in df_name:
        market_input = 'United States'
    elif 'CA' in df_name:
        market_input = 'Canada'
    else:
        raise ValueError(f"Không thể xác định thị trường từ {df_name}")

    # Convert SKU column to string in product_id
    product_id = product_id.iloc[1:, 1:]
    product_id['SKU'] = product_id['SKU'].astype(str)

    # Function to get matching SKU
    def get_matching_sku(campaign_name):
        for sku in product_id['SKU']:
            if sku in campaign_name:
                return sku
        return None

    # Apply the SKU extraction to df
    df['SKU'] = df['Campaign Name'].apply(get_matching_sku)

    # Trim column names
    trim_column_names = lambda x: x.strip()
    df.rename(columns=trim_column_names, inplace=True)

    # Rename columns based on ad_type -> Đổi tên cột
    if ad_type == 'SP':
        df.rename(columns={
            'Click-Thru Rate (CTR)': 'CTR',
            'Cost Per Click (CPC)': 'CPC',
            '7 Day Total Orders (#)': 'Orders',
            'Total Advertising Cost of Sales (ACOS)': 'ACOS',
            'Total Return on Advertising Spend (ROAS)': 'ROAS',
            '7 Day Total Sales': 'Sales'
        }, inplace=True)
    elif ad_type == 'SB':
        df.rename(columns={
            'Click-Thru Rate (CTR)': 'CTR',
            'Cost Per Click (CPC)': 'CPC',
            '14 Day Total Orders (#)': 'Orders',
            'Total Advertising Cost of Sales (ACOS)': 'ACOS',
            'Total Return on Advertising Spend (ROAS)': 'ROAS',
            '14 Day Total Sales': 'Sales'
        }, inplace=True)
    elif ad_type == 'SD':
        df.rename(columns={
            'Click-Thru Rate (CTR)': 'CTR',
            'Cost Per Click (CPC)': 'CPC',
            '14 Day Total Orders (#)': 'Orders',
            'Total Advertising Cost of Sales (ACOS)': 'ACOS',
            'Total Return on Advertising Spend (ROAS)': 'ROAS',
            '14 Day Total Sales': 'Sales'
        }, inplace=True)

    # Clean and convert numerical columns
    for column in ['Budget', 'Spend', 'CPC', 'Sales']:
        if column in df.columns and df[column].dtype == 'object':
            df[column] = df[column].astype(str)
            df[column] = df[column].str.replace(r'[CA|US|\$|,]', '', regex=True)
            df[column] = df[column].astype(float)

    # Additional processing for df
    df['Product Number'] = np.nan
    df['SKU'] = df['SKU'].astype(str)
    df['Cost Type'] = df['Campaign Name'].apply(lambda x: 'CPM' if 'CPM' in x else 'CPC')

    # Campaign name classification
    def classify_campaign_name(campaign_name):
        words = campaign_name.split()
        if ad_type == 'SP':
            if "Auto" in words or "jido" in words:
                return "Auto"
            if "Query" in words or "query" in words:
                return "SP Query"
            if "Research" in words:
                return "Research"
            if "Performance" in words:
                return "Performance"
            if "term" in words or "terms" in words:
                return "search terms"
            if "TD" in words:
                return "TD"
            if "Broad" in words:
                return "SP Broad"
            if "Exact" in words or "EX8" in words:
                return "SP Exact"
            if "TOS" in words:
                return "TOS"
            if "PP" in words:
                return "PP"
            if "PT" in words:
                return "SP PT"
            else:
                return "SP Phrase"
        elif ad_type == 'SB':
            if "Video Ads" in campaign_name:
                if "Phrase" in campaign_name:
                    return "SB Video Phrase"
                if "Broad" in campaign_name:
                    return "SB Video Broad"
                if "Exact" in campaign_name:
                    return "SB Video Exact"
                if "PT" in campaign_name:
                    return "SB Video PT"
                if "Query" in campaign_name:
                    return "SB Video Query"
            return "SB"
        elif ad_type == 'SD':
            return "SD PT"

    df['Campaign Form'] = df['Campaign Name'].apply(classify_campaign_name)

    # Adjustments for Canada market
    if market_input == 'Canada':
        df['Spend'] *= 0.76
        df['Sales'] *= 0.76

    df['Market'] = market_input
    df['Brand'] = brand

    # Additional columns based on ad_type
    if ad_type == 'SP':
        df["Campaign Type"] = 'Sponsored Products'
    elif ad_type == 'SB':
        df['Campaign Type'] = 'Sponsor Brands'
        df['Status'] = np.nan
        df['Budget'] = np.nan
        df['Targeting Type'] = np.nan
        df['Bidding strategy'] = np.nan
    elif ad_type == 'SD':
        df['Campaign Type'] = 'Sponsor Display'
        df['Targeting Type'] = np.nan
        df['Bidding strategy'] = np.nan

    df['Start Date'] = np.nan
    df['End Date'] = np.nan

    # Define the order of columns -> Thêm cột mới tại đây
    order1 = ['Date', 'Campaign Type', 'Campaign Name', 'Bidding strategy',
              'Impressions', 'Clicks', 'Spend', 'Orders', 'Sales',
              'Product Number', 'Campaign Form', 'Market', 'Brand',
              'Cost Type', 'SKU']

    data_merge1 = df[order1]
    return data_merge1

campaigns_no_sku = {'BlueStars': set(), 'Canamax': set()}

for df_name, df in dataframes.items():
    df_cleaned = process_dataframe(df, df_name)
    new_df_name = f"{df_name}_cleaned"
    globals()[new_df_name] = df_cleaned

    brand = df_cleaned['Brand'].iloc[0]
    if 'SKU' in df_cleaned.columns:
        no_sku_campaigns = df_cleaned[df_cleaned['SKU'] == 'None']['Campaign Name'].unique()
        campaigns_no_sku[brand].update(no_sku_campaigns)

ignore_cases = {
    'BlueStars': {'bo di', 'bord', 's','CBB60 Capacitor Auto Catch AlL - new', 'Campaign with presets - B0CDX4BFF6 - 1/4/2025 16:53:52', 'Campaign with presets - B0CGLZ34H5 - 1/4/2025 16:53:52','B0CMCHYG4Y_SD_AUDT_Views Remarketing _Own product  262e14','W10311524 PO5 PO8 AMZ ST','WR57X10032'},
    'Canamax': {'a', 'b', 'c', 'CSR-U2 Video Ads Phrase'}
}

# Define DataFrame names with their corresponding filenames
df_names = {
    'BS_US_SP_link_cleaned': f'BlueStars US SP {start_date_str} - {end_date_str}',
    'BS_US_SB_link_cleaned': f'BlueStars US SB {start_date_str} - {end_date_str}',
    'BS_US_SD_link_cleaned': f'BlueStars US SD {start_date_str} - {end_date_str}',
    'BS_CA_SP_link_cleaned': f'BlueStars CA SP {start_date_str} - {end_date_str}',
    'BS_CA_SB_link_cleaned': f'BlueStars CA SB {start_date_str} - {end_date_str}',
    'BS_CA_SD_link_cleaned': f'BlueStars CA SD {start_date_str} - {end_date_str}'
}

# Combine DataFrames (replace these with your actual DataFrames)
dataframes = {
    'BS_US_SP_link_cleaned': BS_US_SP_link_cleaned,
    'BS_US_SB_link_cleaned': BS_US_SB_link_cleaned,
    'BS_US_SD_link_cleaned': BS_US_SD_link_cleaned,
    'BS_CA_SP_link_cleaned': BS_CA_SP_link_cleaned,
    'BS_CA_SB_link_cleaned': BS_CA_SB_link_cleaned,
    'BS_CA_SD_link_cleaned': BS_CA_SD_link_cleaned
}

df_combined = pd.concat([BS_US_SP_link_cleaned, BS_US_SB_link_cleaned, BS_US_SD_link_cleaned,
                        BS_CA_SP_link_cleaned, BS_CA_SB_link_cleaned, BS_CA_SD_link_cleaned])

df_combined[['Impressions', 'Clicks', 'Spend', 'Orders', 'Sales']] = df_combined[['Impressions', 'Clicks', 'Spend', 'Orders', 'Sales']].fillna(0)
df_combined = df_combined.fillna({'Bidding strategy': "Dynamic bids - down only"})
df_combined = df_combined.drop(columns=['Product Number'])
df_combined['Campaign Type'] = df_combined['Campaign Type'].str.strip()
df_combined['Campaign Name'] = df_combined['Campaign Name'].str.strip()
df_combined['Bidding strategy'] = df_combined['Bidding strategy'].str.strip()
df_combined['Campaign Form'] = df_combined['Campaign Form'].str.strip()
df_combined['Market'] = df_combined['Market'].str.strip()
df_combined['Brand'] = df_combined['Brand'].str.strip()
df_combined['Cost Type'] = df_combined['Cost Type'].str.strip()
df_combined['SKU'] = df_combined['SKU'].str.strip()

# ==================================================================================================
#                                         SETUP EMAIL SENDING FUNCTION
# ==================================================================================================
sender_email = "khangnp.bluestars@gmail.com"
receiver_emails = ["khangnguyenforwork@gmail.com", "duongnt.bluestars@gmail.com", "duybachduybach@gmail.com"]
app_password = "ouia rgwy cuay baoi"

def send_email_with_attachment(subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(receiver_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Thêm file đính kèm nếu có
    if attachment_path:
        try:
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            msg.attach(part)
        except FileNotFoundError:
            print(f"File đính kèm không tồn tại: {attachment_path}")
        except Exception as e:
            print(f"Lỗi khi thêm file đính kèm: {e}")

    try:
        # Kết nối đến máy chủ Gmail và gửi email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)
            server.sendmail(sender_email, receiver_emails, msg.as_string())
        print("Email đã được gửi thành công!")
    except Exception as e:
        print(f"Lỗi khi gửi email: {e}")

# ==================================================================================================
#                                         SEND EMAIL WITH ATTACHMENT OR ZIP
# ==================================================================================================
send_zip = True
missing_sku_details = []

for brand, campaigns in campaigns_no_sku.items():
    filtered_campaigns = campaigns - ignore_cases.get(brand, set())

    if filtered_campaigns:
        detail = f"Brand {brand} có các campaigns sau chưa có SKU cập nhật:\n"
        for campaign in sorted(filtered_campaigns):
            detail += f"- {campaign}\n"
        missing_sku_details.append(detail)
        send_zip = False

if missing_sku_details:
    subject = "MISSING SKU FOR MULTIPLE BRANDS"
    body = "Các brands sau có campaigns chưa được cập nhật SKU:\n\n" + "\n\n".join(missing_sku_details)
    send_email_with_attachment(subject, body, None)
else:
    zip_file_name = f"Weekly Marketing Data {start_date_str} - {end_date_str}.zip"
    with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for df_key, df in dataframes.items():
            file_name = f"{df_names[df_key]}.xlsx"
            with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            zipf.write(file_name)
            os.remove(file_name)

    subject = f"Marketing Weekly Data Report {start_date_str} - {end_date_str}"
    body = f"Marketing Weekly Data Report {start_date_str} - {end_date_str}"
    send_email_with_attachment(subject, body, zip_file_name)

    os.remove(zip_file_name)















