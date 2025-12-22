from __future__ import annotations

import csv
import html
import os
import time
from configparser import ConfigParser
from datetime import datetime
from pathlib import Path
from typing import Optional, List

from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from tkinter import Tk, Toplevel, Label, StringVar, messagebox
from tkinter import ttk as tkttk


# ─── Configuration ─────────────────────────────────────────────
CONFIG_FILE_PATH = Path(r"C:\Users\nsikder\Downloads\config.ini")
LOG_HTML_FILE = Path("validation_log.html")

# ─── Global log store ──────────────────────────────────────────
log_entries: list[str] = []
html_report_written: bool = False


# ─── Logging Utilities ─────────────────────────────────────────
def log(message: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{ts}] {message}"
    log_entries.append(entry)
    print(entry)


# ─── HTML Report ───────────────────────────────────────────────
def save_logs_to_html(
    customer: str,
    export_type: str,
    filename: Path = LOG_HTML_FILE,
    sample_table_html: Optional[str] = None,
) -> None:
    global html_report_written

    try:
        with filename.open("w", encoding="utf-8") as f:
            f.write("<!doctype html><html><head><meta charset='utf-8'>")
            f.write("<title>Export Validation Log</title>")
            f.write("""
            <style>
                body{font-family:Segoe UI,Arial;background:#f6f7f9;padding:20px}
                .card{background:#fff;padding:20px;border-radius:8px}
                .entry{background:#f3f4f6;padding:8px;margin:6px 0;border-radius:6px}
                h1{color:#2f6f17}
            </style>
            </head><body>
            """)

            f.write("<div class='card'>")
            f.write("<h1>Export Dashboard Validation Log</h1>")
            f.write(f"<p><b>Customer:</b> {html.escape(customer)}</p>")
            f.write(f"<p><b>Export:</b> {html.escape(export_type)}</p>")
            f.write("<hr>")

            for e in log_entries:
                f.write(f"<div class='entry'>{html.escape(e)}</div>")

            if sample_table_html:
                f.write(sample_table_html)

            f.write("</div></body></html>")

        html_report_written = True
        log(f"HTML report saved: {filename.resolve()}")

    except Exception as e:
        log(f"Failed to save HTML log: {e}")


# ─── Config & Driver Setup ─────────────────────────────────────
class ConfParser:
    def __init__(self, config_file_path: Path) -> None:
        self.config = ConfigParser()
        if not config_file_path.exists():
            raise FileNotFoundError(f"Config not found: {config_file_path}")
        self.config.read(config_file_path)
        log(f"Config loaded: {config_file_path}")


class ChromeDriverSetup(ConfParser):
    def __init__(self, config_file_path: Path):
        super().__init__(config_file_path)

        options = webdriver.ChromeOptions()

        try:
            options.add_argument(self.config["path"]["chrome_profile"])
        except Exception:
            pass

        prefs = {"safebrowsing.enabled": True}
        options.add_experimental_option("prefs", prefs)

        service = Service(self.config["path"]["chrome_driver"])
        self.driver = webdriver.Chrome(service=service, options=options)

        log("Chrome driver initialized")


# ─── Helper Functions ──────────────────────────────────────────
class SupportiveFunctions:
    driver: webdriver.Chrome

    def ajax_preloader_wait(self, timeout: int = 300) -> None:
        try:
            time.sleep(1)
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element(
                    (By.XPATH, "//div[contains(@class,'ajax_preloader')]")
                )
            )
        except Exception:
            log("Ajax preloader wait skipped or timed out")


# ─── Progress Window ───────────────────────────────────────────
class ProgressWindow:
    def __init__(self, master: Tk, total_steps: int) -> None:
        self.window = Toplevel(master)
        self.window.title("Validation Progress")
        self.window.geometry("700x180")
        self.window.configure(bg="#4f8611")
        self.window.resizable(False, False)

        self.total_steps = max(1, total_steps)
        self.step = 0

        self.status_var = StringVar(value="Starting...")
        Label(
            self.window,
            textvariable=self.status_var,
            bg="#4f8611",
            fg="white",
            font=("Arial", 13, "bold"),
            wraplength=640,
            justify="center",
        ).pack(pady=10)

        self.progress = tkttk.Progressbar(
            self.window,
            orient="horizontal",
            length=560,
            mode="determinate",
            maximum=self.total_steps,
        )
        self.progress.pack(pady=6)

        self.percent = Label(
            self.window,
            text="0%",
            bg="#4f8611",
            fg="white",
            font=("Arial", 11, "bold"),
        )
        self.percent.pack()

        self.window.update()

    def update(self, message: str) -> None:
        self.step = min(self.step + 1, self.total_steps)
        self.status_var.set(message)
        self.progress["value"] = self.step
        self.percent.config(text=f"{int((self.step / self.total_steps) * 100)}%")
        log(message)
        self.window.update_idletasks()
        time.sleep(0.3)

    def complete(self) -> None:
        self.status_var.set("✅ Validation completed")
        self.progress["value"] = self.total_steps
        self.percent.config(text="100%")
        log("Validation completed")
        self.window.update_idletasks()
        time.sleep(1)
        self.window.destroy()


# ─── Login + Validation Runner ─────────────────────────────────
class CozevaLogin(ChromeDriverSetup, SupportiveFunctions):
    def certlogin_cozeva(self, customer: str, progress: ProgressWindow) -> None:
        """Perform login to CERT and select customer via UI interactions."""
        try:
            progress.update("Logging into Cozeva (CERT)...")
            self.driver.get(self.config.get("cert", "logout_url", fallback="about:blank"))
            self.driver.get(self.config.get("cert", "login_url", fallback="about:blank"))
            self.driver.maximize_window()

            user = os.environ.get("CS2User")
            pwd = os.environ.get("CS2Password")
            if not all((user, pwd)):
                raise RuntimeError("Environment variables CS2User / CS2Password not set.")

            self.driver.find_element(By.ID, "edit-name").send_keys(user)
            self.driver.find_element(By.ID, "edit-pass").send_keys(pwd)
            self.driver.find_element(By.ID, "edit-submit").click()

            WebDriverWait(self.driver, 120).until(EC.presence_of_element_located((By.ID, "reason_textbox")))
            self.driver.find_element(By.XPATH, "//*[@id='select-customer']").click()
            self.driver.find_element(By.XPATH, f"//*[contains(text(), '{str(customer)}')]").click()

            WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "reason_textbox")))
            reason_text = self.config.get("credentials", "user_search_reason", fallback="")
            self.driver.find_element(By.ID, "reason_textbox").send_keys(reason_text)
            self.driver.find_element(By.ID, "edit-submit").click()
            time.sleep(5)

            self.ajax_preloader_wait()
            progress.update("Logged in successfully (CERT).")
        except Exception as e:
            log(f"❌ Login Error (CERT): {e}")
            messagebox.showerror("Login Error", str(e))
            raise

    def prodlogin_cozeva(self, customer: str, progress: ProgressWindow) -> None:
        """Perform login to PROD and select customer via UI interactions."""
        try:
            progress.update("Logging into Cozeva (PROD)...")
            self.driver.get(self.config.get("prod", "logout_url", fallback="about:blank"))
            self.driver.get(self.config.get("prod", "login_url", fallback="about:blank"))
            self.driver.maximize_window()

            user = os.environ.get("CS2User")
            pwd = os.environ.get("CS2Password")
            if not all((user, pwd)):
                raise RuntimeError("Environment variables CS2User / CS2Password not set.")

            self.driver.find_element(By.ID, "edit-name").send_keys(user)
            self.driver.find_element(By.ID, "edit-pass").send_keys(pwd)
            self.driver.find_element(By.ID, "edit-submit").click()

            WebDriverWait(self.driver, 120).until(EC.presence_of_element_located((By.ID, "reason_textbox")))
            self.driver.find_element(By.XPATH, "//*[@id='select-customer']").click()
            self.driver.find_element(By.XPATH, f"//*[contains(text(), '{str(customer)}')]").click()

            WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "reason_textbox")))
            reason_text = self.config.get("credentials", "user_search_reason", fallback="")
            self.driver.find_element(By.ID, "reason_textbox").send_keys(reason_text)
            self.driver.find_element(By.ID, "edit-submit").click()
            time.sleep(5)

            self.ajax_preloader_wait()
            progress.update("Logged in successfully (PROD).")
        except Exception as e:
            log(f"❌ Login Error (PROD): {e}")
            messagebox.showerror("Login Error", str(e))
            raise

    def logout(self, progress: ProgressWindow, customer: str) -> None:
        try:
            self.driver.quit()
        except Exception:
            pass
        progress.complete()
        save_logs_to_html(customer, "User Search Validation")


class user_search(ChromeDriverSetup, SupportiveFunctions):
    def __init__(self, driver, config):
        self.driver = driver
        self.config = config

    def users_list(self, progress: ProgressWindow):
        try:
            progress.update("Opening User List page...")
            self.driver.get(self.config.get("user_list", "list_url", fallback="about:blank"))
            self.ajax_preloader_wait()
            time.sleep(3)
            self.driver.find_element(By.XPATH, "//a[@data-target='table_dropdown_people_list']").click()
            self.driver.find_element(By.XPATH, "(//input[starts-with(@id,'select-dropdown-')])[1]").click()

        except Exception as e:
            log(f"❌ users list didn't open: {e}")
            raise

        print("Okay")


# ─── ENTRY POINT CALLED FROM UI ─────────────────────────────────
def run_user_validation(
    master_window: Tk,
    customer: str,
    selected_areas: List[str],
    environment: str = "CERT",
) -> None:
    """
    SAFE entry point for Tkinter SUBMIT button.
    """

    progress = ProgressWindow(master_window, total_steps=3 + len(selected_areas))

    try:
        runner = CozevaLogin(CONFIG_FILE_PATH)
        # 1️⃣ LOGIN
        if environment.upper() == "CERT":
            runner.certlogin_cozeva(customer, progress)
        else:
            raise NotImplementedError("PROD login not wired yet")
        # 2️⃣ OPEN USER LIST
        search = user_search(runner.driver, runner.config)
        search.users_list(progress)

        # Area validations (placeholder hooks)
        for area in selected_areas:
            progress.update(f"Validating: {area}")
            time.sleep(0.4)  # replace with real Selenium logic per area

        runner.logout(progress, customer)

    except Exception as e:
        log(f"Validation failed: {e}")
        messagebox.showerror("Validation Error", str(e))
        try:
            progress.window.destroy()
        except Exception:
            pass
