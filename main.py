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

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Initialize Selenium WebDriver for Edge
def initialize_driver():
    edge_options = Options()
    edge_options.add_argument("--disable-logging")  # Disable logging

    # Path to your msedgedriver.exe (Update this with your actual path)
    edge_driver_path = r"C:\Users\rsain\Downloads\edgedriver_win64\msedgedriver.exe"  # Update this with your path

    # Use the Service class to specify the driver path
    service = Service(edge_driver_path)
    driver = webdriver.Edge(service=service, options=edge_options)  # Use Service to provide the driver
    logging.info("Edge WebDriver initialized.")
    return driver

# Function to handle the login process
def login():
    username = username_entry.get()
    password = password_entry.get()

    if not username or not password:
        messagebox.showerror("Error", "Please enter both username and password.")
        logging.error("Username or password is missing.")
        return

    driver = None  # Initialize the driver variable outside the try block

    try:
        logging.info(f"Starting login process for user: {username}")
        driver = initialize_driver()

        # Open MaaS360 login page
        driver.get("https://login.maas360.com")
        logging.info("Navigated to MaaS360 login page.")

        # Wait for the page to load
        wait = WebDriverWait(driver, 20)  # Increased wait time

        # Fill in the username
        username_field = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/section/div/section/form/div/input')))
        username_field.send_keys(username)
        logging.info(f"Entered username: {username}")

        # Click the "Continue" button
        continue_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/section/div/section/form/div/div[2]/input')))
        continue_button.click()
        logging.info("Clicked 'Continue' button.")

        # Fill in the password
        password_field = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/form/div[2]/div/div[1]/div/div/input')))
        password_field.send_keys(password)
        logging.info("Entered password.")

        # Click the "Login" button
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/form/div[3]/div/div/button')))
        login_button.click()
        logging.info("Clicked 'Login' button.")

        # Wait for the dashboard to load
        wait.until(EC.url_to_be("https://m3.maas360.com/emc/?"))
        logging.info("Logged in successfully, dashboard loaded.")

        # Click the "Users" menu item
        try:
            users_menu_item = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/ul[1]/li[3]')))
            logging.info("Found 'Users' menu item, now clicking.")
            users_menu_item.click()
            logging.info("Clicked 'Users' menu item.")
        except Exception as e:
            logging.error(f"Error clicking 'Users' menu item: {e}")
            messagebox.showerror("Error", f"Error clicking 'Users' menu item: {e}")
            return

        # Add a delay to ensure everything has fully loaded before trying to click
        logging.info("Waiting for 6 seconds before attempting to click 'Add User'.")
        time.sleep(2)  # Sleep for 6 seconds to allow the page to finish loading

        # Wait for the iframe containing the 'Add User' button to load
        iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='Content']")))
        driver.switch_to.frame(iframe)  # Switch to the iframe containing the button
        logging.info("Switched to iframe containing 'Add User' button.")

        # Wait for the "Add User" button to be clickable using the updated XPath
        try:
            add_user_button = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='addUserButton']")))  # Updated XPath
            logging.info("Found 'Add User' button, executing JavaScript to trigger the modal.")
            driver.execute_script("arguments[0].click();", add_user_button)  # Trigger the click event using JS
            logging.info("Clicked 'Add User' button using JavaScript.")
        except Exception as e:
            logging.error(f"Error clicking 'Add User' button using JavaScript: {e}")
            messagebox.showerror("Error", f"Error clicking 'Add User' button: {e}")
            return

        # Wait for the 'New Hire Name' field to be visible and interactable
        logging.info("Waiting for 'New Hire Name' field.")
        full_name_field = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='addUserBasicForm.fullName']")))

        # Now fill in the "New Hire Name" field
        new_hire_name = new_hire_name_entry.get()  # Get the value from the Tkinter entry field
        logging.info(f"Filling 'New Hire Name' with: {new_hire_name}")
        
        # Fill the field
        full_name_field.clear()  # Clear any pre-existing text (just in case)
        full_name_field.send_keys(new_hire_name)

        # Fill out the "Username" field (first initial + last name, lowercase)
        username_value = new_hire_name.strip().lower().replace(" ", "")
        username_field = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='addUserBasicForm.username']")))

        username_field.clear()
        username_field.send_keys(username_value)
        logging.info(f"Filling 'Username' with: {username_value}")

        # Fill out the "Password" field
        password_field = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='addUserBasicForm.userPasswordHidden']")))

        password_field.clear()
        password_field.send_keys("w2LuC#Ad")
        logging.info("Filling 'Password' with: w2LuC#Ad")

        # Fill out the "Domain" field
        domain_field = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='addUserBasicForm.domain']")))

        domain_field.clear()
        domain_field.send_keys("mesfire")
        logging.info("Filling 'Domain' with: mesfire")

        # Fill out the "Email" field (username + "@mdm.mesfire.com")
        email_field = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='addUserBasicForm.notificationModel.userEmail']")))

        email_field.clear()
        email_field.send_keys(f"{username_value}@mdm.mesfire.com")
        logging.info(f"Filling 'Email' with: {username_value}@mdm.mesfire.com")

        # Now click the 'Save' button
        save_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/form/div[3]/span/span/input[2]")))
        save_button.click()
        logging.info("Clicked 'Save' button.")


        # Wait for the password modal to appear
        logging.info("Waiting for the password modal to appear.")
        password_modal = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='ui-dialog-content']//input[@id='centralisedPassword']")))
        logging.info("Password modal appeared.")

        # Enter the password in the input field
        password_field = driver.find_element(By.XPATH, "//div[@class='ui-dialog-content']//input[@id='centralisedPassword']")
        password_field.send_keys(password)  # Use the password from the Tkinter form
        logging.info("Entered password in the modal.")

        # Wait for the "Confirm" button to become enabled
        confirm_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='ui-dialog-content']//input[@id='pwdSubmit']")))

        # Click the "Confirm" button
        confirm_button.click()
        logging.info("Clicked 'Confirm' button after entering the password.")


        # Wait for an additional 5 seconds before closing the driver
        logging.info("Waiting for 5 seconds before closing the driver.")
        time.sleep(5)  # Wait 5 seconds before closing the browser

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

    finally:
        if driver:
            # Keep the browser open until manually closed
            logging.info("Driver open. Close manually when done.")
            driver.quit()
            logging.info("Driver closed.")


# Create the main application window
app = tk.Tk()

# Set window size and position
app.geometry("400x400")
app.resizable(False, False)

# Add a label and entry for the username
username_label = tk.Label(app, text="Username:")
username_label.pack(pady=10)
username_entry = tk.Entry(app, width=40)
username_entry.pack()

# Add a label and entry for the password
password_label = tk.Label(app, text="Password:")
password_label.pack(pady=10)
password_entry = tk.Entry(app, width=40, show="*")  # Hide password input
password_entry.pack()

# Add a label and entry for New Hire Name
new_hire_name_label = tk.Label(app, text="New Hire Name:")
new_hire_name_label.pack(pady=10)
new_hire_name_entry = tk.Entry(app, width=40)  # New entry field for New Hire Name
new_hire_name_entry.pack()

# Add a login button
login_button = tk.Button(app, text="Login", command=login)
login_button.pack(pady=20)

# Run the application
app.mainloop()
