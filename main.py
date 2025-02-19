import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
import time
import logging
import os
import csv

# Set the base folder for logs and CSV files
BASE_FOLDER = r"C:\maas360_users"
if not os.path.exists(BASE_FOLDER):
    os.makedirs(BASE_FOLDER)

# Configure initial logging (will be reconfigured per run)
logging.basicConfig(
    filename=os.path.join(BASE_FOLDER, "default.log"),
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Initialize Selenium WebDriver for Edge
def initialize_driver():
    edge_options = Options()
    edge_options.add_argument("--disable-logging")
    edge_driver_path = r"C:\Users\rsain\Downloads\edgedriver_win64\msedgedriver.exe"  # Update with your path
    service = Service(edge_driver_path)
    driver = webdriver.Edge(service=service, options=edge_options)
    logging.info("Edge WebDriver initialized.")
    return driver

# Function to handle the password modal dialog (optimized for speed)
def handle_password_modal(driver, wait, password):
    try:
        # Wait for the modal dialog to be visible.
        modal = wait.until(EC.visibility_of_element_located(
            (By.XPATH, "//div[contains(@class, 'ui-dialog-content')]")))
        logging.info("Modal is visible.")
        
        # If the modal might be within an iframe, try to switch into it.
        try:
            modal_iframe = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//iframe[contains(@src, 'password') or contains(@id, 'password')]")))
            driver.switch_to.frame(modal_iframe)
            logging.info("Switched to modal's iframe.")
        except Exception:
            logging.info("Modal does not appear to be inside an iframe. Continuing in current context.")
        
        # Locate the password input field using its correct ID "centralisedPassword"
        password_input = wait.until(EC.element_to_be_clickable((By.ID, "centralisedPassword")))
        password_input.clear()
        # Set the password value directly for speed
        driver.execute_script("arguments[0].value = arguments[1];", password_input, password)
        logging.info("Set password in modal field using id 'centralisedPassword'.")
        
        # Dispatch events to trigger any JavaScript listeners.
        driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", password_input)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", password_input)
        
        # Reduced wait time for the confirm button to update.
        time.sleep(0.2)
        
        # Locate the confirm button using its ID "pwdSubmit" and enable it.
        confirm_button = wait.until(EC.presence_of_element_located((By.ID, "pwdSubmit")))
        driver.execute_script("arguments[0].removeAttribute('disabled');", confirm_button)
        driver.execute_script("arguments[0].classList.remove('disabled-btn');", confirm_button)
        confirm_button = wait.until(EC.element_to_be_clickable((By.ID, "pwdSubmit")))
        confirm_button.click()
        logging.info("Clicked 'Confirm' button in modal.")
        
        # Switch back to default content.
        driver.switch_to.default_content()
    except Exception as e:
        logging.error(f"Error handling password modal: {e}")
        raise

# Global driver variable
driver = None

# Function to handle the login and user creation process
def login():
    global driver
    username_val = username_entry.get()
    password_val = password_entry.get()
    new_hire_name = new_hire_name_entry.get().strip()
    
    if not username_val or not password_val or not new_hire_name:
        messagebox.showerror("Error", "Please enter username, password, and new hire name.")
        logging.error("Missing one or more required fields.")
        return

    # Create dynamic log file based on new hire name (firstnamelastname.log)
    log_filename = new_hire_name.replace(" ", "").lower() + ".log"
    log_path = os.path.join(BASE_FOLDER, log_filename)
    # Remove existing handlers and reconfigure logging for this run.
    logger = logging.getLogger()
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    file_handler = logging.FileHandler(log_path)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.setLevel(logging.INFO)
    logging.info(f"Logging started for new hire: {new_hire_name} into file {log_path}")

    # Prepare CSV file path (firstnamelastname.csv)
    csv_filename = new_hire_name.replace(" ", "").lower() + ".csv"
    csv_path = os.path.join(BASE_FOLDER, csv_filename)

    # Prepare formatted name for CSV (Lastname, Firstname)
    name_parts = new_hire_name.lower().split()
    if len(name_parts) >= 2:
        formatted_name = f"{name_parts[-1].capitalize()}, {name_parts[0].capitalize()}"
    else:
        formatted_name = new_hire_name.title()

    try:
        logging.info(f"Starting login process for user: {username_val}")
        driver = initialize_driver()
        driver.get("https://login.maas360.com")
        logging.info("Navigated to MaaS360 login page.")
        wait = WebDriverWait(driver, 20)

        # Fill in the login username
        username_field = wait.until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div/div/section/div/section/form/div/input')))
        username_field.send_keys(username_val)
        logging.info(f"Entered username: {username_val}")

        # Click "Continue"
        continue_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[1]/div/div/section/div/section/form/div/div[2]/input')))
        continue_button.click()
        logging.info("Clicked 'Continue' button.")

        # Fill in the login password.
        login_password_field = wait.until(EC.presence_of_element_located(
            (By.XPATH, '/html/body/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/form/div[2]/div/div[1]/div/div/input')))
        login_password_field.send_keys(password_val)
        logging.info("Entered login password.")

        # Click "Login"
        login_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/form/div[3]/div/div/button')))
        login_button.click()
        logging.info("Clicked 'Login' button.")

        # Wait for dashboard to load.
        wait.until(EC.url_to_be("https://m3.maas360.com/emc/?"))
        logging.info("Logged in successfully, dashboard loaded.")

        # Click "Users" menu item.
        try:
            users_menu_item = wait.until(EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[2]/div[3]/ul[1]/li[3]')))
            users_menu_item.click()
            logging.info("Clicked 'Users' menu item.")
        except Exception as e:
            logging.error(f"Error clicking 'Users' menu item: {e}")
            messagebox.showerror("Error", f"Error clicking 'Users' menu item: {e}")
            return

        logging.info("Waiting for 2 seconds before attempting to click 'Add User'.")
        time.sleep(2)

        # Switch to the iframe containing the "Add User" button.
        iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='Content']")))
        driver.switch_to.frame(iframe)
        logging.info("Switched to iframe containing 'Add User' button.")

        # Click "Add User" button via JavaScript.
        try:
            add_user_button = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//*[@id='addUserButton']")))
            driver.execute_script("arguments[0].click();", add_user_button)
            logging.info("Clicked 'Add User' button using JavaScript.")
        except Exception as e:
            logging.error(f"Error clicking 'Add User' button: {e}")
            messagebox.showerror("Error", f"Error clicking 'Add User' button: {e}")
            return

        # Wait for "New Hire Name" field and fill it in.
        logging.info("Waiting for 'New Hire Name' field.")
        full_name_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='addUserBasicForm.fullName']")))
        full_name_field.clear()
        full_name_field.send_keys(new_hire_name)
        logging.info(f"Filling 'New Hire Name' with: {new_hire_name}")

        # Fill "Username" for the new user.
        if len(name_parts) >= 2:
            username_value = f"{name_parts[0]}.{name_parts[-1]}"
        else:
            username_value = new_hire_name.lower()
        username_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='addUserBasicForm.username']")))
        username_field.clear()
        username_field.send_keys(username_value)
        logging.info(f"Filling 'Username' with: {username_value}")

        # Fill "Password" for the new user.
        user_password_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='addUserBasicForm.userPasswordHidden']")))
        user_password_field.clear()
        user_password_field.send_keys("w2LuC#Ad")
        logging.info("Filling 'Password' with: w2LuC#Ad")

        # Fill "Domain".
        domain_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='addUserBasicForm.domain']")))
        domain_field.clear()
        domain_field.send_keys("mesfire")
        logging.info("Filling 'Domain' with: mesfire")

        # Fill "Email".
        email_field = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='addUserBasicForm.notificationModel.userEmail']")))
        email_field.clear()
        email_field.send_keys(f"{username_value}@mdm.mesfire.com")
        logging.info(f"Filling 'Email' with: {username_value}@mdm.mesfire.com")

        # Click "Save".
        save_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "/html/body/div[2]/div[2]/div/form/div[3]/span/span/input[2]")))
        save_button.click()
        logging.info("Clicked 'Save' button.")

        # Switch back to default content so the modal is visible.
        driver.switch_to.default_content()
        logging.info("Switched back to default content after clicking 'Save'.")

        # Handle the password modal.
        handle_password_modal(driver, wait, password_val)

        # Append CSV record.
        email_value = f"{username_value}@mdm.mesfire.com"
        with open(csv_path, mode='a', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([username_value, formatted_name, "w2LuC#Ad", email_value])
        logging.info(f"CSV record written to {csv_path}.")

        # Show success message and clear the form fields.
        messagebox.showinfo("Success", f"{new_hire_name} has been added to MaaS360.")
        username_entry.delete(0, tk.END)
        password_entry.delete(0, tk.END)
        new_hire_name_entry.delete(0, tk.END)
        logging.info("Cleared the form for a new entry.")

        # Optional wait before ending this run.
        time.sleep(5)

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

    finally:
        if driver:
            driver.quit()
            logging.info("Driver closed.")

# Create the main application window
app = tk.Tk()
app.title("MaaS360 User Creation")
app.geometry("400x400")
app.resizable(False, False)

username_label = tk.Label(app, text="Username:")
username_label.pack(pady=10)
username_entry = tk.Entry(app, width=40)
username_entry.pack()

password_label = tk.Label(app, text="Password:")
password_label.pack(pady=10)
password_entry = tk.Entry(app, width=40, show="*")
password_entry.pack()

new_hire_name_label = tk.Label(app, text="New Hire Name:")
new_hire_name_label.pack(pady=10)
new_hire_name_entry = tk.Entry(app, width=40)
new_hire_name_entry.pack()

login_button = tk.Button(app, text="Login", command=login)
login_button.pack(pady=20)

app.mainloop()
