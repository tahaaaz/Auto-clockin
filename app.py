import time
import re
import imaplib
import email
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from PIL import Image
import pytesseract
import sys



from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import tempfile

chrome_options = Options()
chrome_options.add_argument("--headless=new")  # run in headless mode
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Use a temporary user-data-dir to avoid conflicts
temp_dir = tempfile.mkdtemp()
chrome_options.add_argument(f"--user-data-dir={temp_dir}")

driver = webdriver.Chrome(options=chrome_options)



# Force UTF-8 encoding for Windows console
sys.stdout.reconfigure(encoding='utf-8')

# ================== CONFIG ==================
KEKA_EMAIL = "tahazuberi@lambdatest.com"
KEKA_PASSWORD = "Keka@1234"

GMAIL_EMAIL = "tahazuberi@lambdatest.com"
GMAIL_APP_PASSWORD = "zfzexumndqnqumhe"
# ============================================

# Path to Tesseract OCR executable
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def get_otp():
    """Fetch latest OTP from Gmail inbox"""
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_EMAIL, GMAIL_APP_PASSWORD)
    mail.select("inbox")
    status, messages = mail.search(None, "ALL")
    if status != "OK":
        return None
    email_ids = messages[0].split()
    for eid in email_ids[::-1][:10]:  # last 10 emails
        _, data = mail.fetch(eid, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)
        content = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    content += part.get_payload(decode=True).decode(errors="ignore")
        else:
            content = msg.get_payload(decode=True).decode(errors="ignore")
        match = re.search(r"OTP:\s*(\d{6})", content)
        if match:
            otp = match.group(1)
            mail.store(eid, '+FLAGS', '\\Seen')
            return otp
    return None

def read_captcha(file_path):
    """Read captcha image with pytesseract"""
    return pytesseract.image_to_string(Image.open(file_path)).strip().upper()

def main():
    driver = webdriver.Chrome()
    driver.maximize_window()
    wait = WebDriverWait(driver, 20)

    # Step 1: Open login page
    driver.get("https://app.keka.com/Account/Login")

    # Enter email
    driver.find_element(By.ID, "email").send_keys(KEKA_EMAIL)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/div[1]/div[2]/form[1]/div/button").click()

    # Switch to password login
    driver.find_element(By.XPATH, "/html/body/div/div[2]/div[1]/div[2]/div[2]/div[2]/button/div/p").click()

    MAX_RETRIES = 10
    login_success = False

    for attempt in range(1, MAX_RETRIES+1):
        print(f"Attempt {attempt} of {MAX_RETRIES}")

        # Enter password
        password_field = wait.until(EC.element_to_be_clickable((By.ID, "password")))
        password_field.clear()
        password_field.send_keys(KEKA_PASSWORD)
        time.sleep(0.5)

        # Wait for captcha image
        try:
            captcha_img = wait.until(EC.visibility_of_element_located((By.ID, "imgCaptcha")))
        except TimeoutException:
            print("X Captcha not found, retrying...")
            continue

        # Extra wait to ensure captcha renders
        retries = 0
        while (captcha_img.size['width'] == 0 or captcha_img.size['height'] == 0) and retries < 5:
            time.sleep(0.5)
            captcha_img = driver.find_element(By.ID, "imgCaptcha")
            retries += 1

        # Save captcha screenshot
        captcha_path = os.path.join(os.getcwd(), "captcha.png")
        captcha_img.screenshot(captcha_path)
        captcha_text = read_captcha(captcha_path)
        print(f"Predicted Captcha: {captcha_text}")

        # Enter captcha
        captcha_field = driver.find_element(By.ID, "captcha")
        captcha_field.click()
        captcha_field.clear()
        for c in captcha_text:
            captcha_field.send_keys(c)
            time.sleep(0.15)

        # Click login
        driver.find_element(By.XPATH, "/html/body/div/div[2]/div[1]/div[2]/form/div/button").click()
        time.sleep(3)

        # ✅ After successful captcha - click "Send code to Gmail"
        try:
            otp_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div[1]/div[2]/div/form[2]/button"))
            )
            otp_button.click()
            print("✅ Clicked 'Send code to Gmail' button successfully!")
            login_success = True
            break
        except TimeoutException:
            print("X Captcha failed, refreshing...")
            try:
                driver.find_element(By.ID, "btnRefreshCaptcha").click()
                time.sleep(2)
            except:
                pass

    if not login_success:
        print("🚨 Captcha failed after maximum attempts.")
        driver.quit()
        return

    # Step 2: Fetch OTP automatically
    otp_val = None
    for _ in range(6):
        otp_val = get_otp()
        if otp_val:
            break
        print("Waiting for OTP email...")
        time.sleep(5)

    if not otp_val:
        print("❌ OTP not found!")
        driver.quit()
        return

    print(f"OTP received: {otp_val}")

    # Enter OTP
    otp_input = wait.until(EC.presence_of_element_located((By.ID, "code")))
    otp_input.send_keys(otp_val)
    driver.find_element(By.XPATH, "/html/body/div/div[2]/div[1]/div[2]/form/div/button").click()

    # Wait for dashboard
    wait.until(EC.presence_of_element_located((By.ID, "accordion")))
    print("✅ Logged in successfully!")

    # Navigate to Attendance
    me_button = driver.find_element(By.XPATH, "//*[@id='accordion']/li[2]/a/span[2]")
    ActionChains(driver).move_to_element(me_button).perform()
    attendance_xpath = "//*[@id='accordion']/li[2]/div/ul/li[1]/a/span"
    wait.until(EC.visibility_of_element_located((By.XPATH, attendance_xpath)))
    driver.find_element(By.XPATH, attendance_xpath).click()
    print("📅 Clicked Attendance...")

    # Handle notifications
    try:
        notif_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='preload']/xhr-app-root/xhr-page-header/notification-prompt/div/div/div[2]/button[1]"))
        )
        notif_btn.click()
        print("🔕 Notification prompt closed.")
    except:
        print("No notification prompt found.")

    # Final clock-in
    clockin_xpaths = [
        "//*[@id='preload']/xhr-app-root/div/employee-me/div/employee-attendance/div/div/div/div/employee-attendance-stats/div/div[3]/employee-attendance-request-actions/div/div/div/div/div[2]/div/div[1]/a[1]",
        "//*[@id='preload']/xhr-app-root/div/employee-me/div/employee-attendance/div/div/div/div/employee-attendance-stats/div/div[3]/employee-attendance-request-actions/div/div/div/div/div[2]/div/div[1]/div[1]/button",
        "//*[@id='preload']/xhr-app-root/div/employee-me/div/employee-attendance/div/div/div/div/employee-attendance-stats/div/div[3]/employee-attendance-request-actions/div/div/div/div/div[2]/div/div[1]/div[1]/button[1]"
    ]

    for xpath in clockin_xpaths:
        try:
            clockin_btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            clockin_btn.click()
            print("🕒 Clocked in successfully!")
            break
        except:
            continue

    time.sleep(5)
    driver.quit()


if __name__ == "__main__":
    main()
