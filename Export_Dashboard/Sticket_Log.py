from __future__ import annotations
import csv
import html
import os
import time
from configparser import ConfigParser
from datetime import datetime
from pathlib import Path
from typing import Optional
from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from tkinter import Tk, Toplevel, Label, StringVar, messagebox
from tkinter import ttk as tkttk

# ─── Configuration ─────────────────────────────────────────────
CONFIG_FILE_PATH = Path(r"C:\python projects\config.ini")
DOWNLOAD_DIR = Path(r"C:\python projects\Export_Dashboard\Exported Files")
LOG_HTML_FILE = Path("sticket_validation_log.html")

# ─── Global log store ─────────────────────────────────────────────
log_entries: list[str] = []
html_report_written: bool = False  # avoid overwriting report once written


def log(message: str) -> None:
    """Append message to global log (timestamped) and also print to stdout."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{ts}] {message}"
    log_entries.append(entry)
    print(entry)


def extract_failed_logs(logs: list[str]) -> list[str]:
    return [entry for entry in logs if "❌" in entry]


def extract_passed_logs(logs: list[str]) -> list[str]:
    return [entry for entry in logs if "✅" in entry]


def save_logs_to_html(
    selected_customer: str,
    patient_id: str,
    filename: Path = LOG_HTML_FILE,
    sample_table_html: Optional[str] = None,
) -> None:
    """
    Write log_entries to a styled HTML file.
    Adds:
    - Failed Cases Summary at the top (after divider)
    - Download PDF + Print buttons
    - Optional CSV sample table with match coloring
    """
    global html_report_written
    try:
        outpath = Path(filename)
        failed_logs = extract_failed_logs(log_entries)
        passed_logs = extract_passed_logs(log_entries)

        with outpath.open("w", encoding="utf-8") as f:
            f.write("<!doctype html><html><head><meta charset='utf-8'>")
            f.write("<title>Export Validation Log</title>")

            f.write(
                """
            <style>
                :root{--brand:#2f6f17;--muted:#4a5568;--bg:#f6f7f9}
                body{font-family:Segoe UI, Arial, sans-serif;background:var(--bg);padding:20px;margin:0}
                .page-wrap{max-width:1100px;margin:20px auto;position:relative}
                .card{background:#fff;border-radius:8px;padding:22px 22px 30px 22px;
                      box-shadow:0 2px 6px rgba(0,0,0,0.08);position:relative}
                h1{color:var(--brand);margin:0 0 10px 0}
                h2{margin:0 0 8px 0}
                .meta{color:#555;font-size:0.95rem;margin-bottom:12px}
                .entry{font-family:Times New Roman;background:#f3f4f6;
                       padding:8px;margin:6px 0;border-radius:6px}
                table.sample{border-collapse:collapse;width:100%;margin-top:12px}
                table.sample th, table.sample td{
                    border:1px solid #e9ecef;padding:6px;text-align:left}
                table.sample thead th{background:#f3f4f6}

                .top-right-actions{
                    position:absolute;top:12px;right:12px;display:flex;gap:8px;z-index:10}
                .btn-download{
                    padding:8px 12px;border-radius:6px;background:var(--brand);
                    color:white;border:none;cursor:pointer;font-weight:600}
                .btn-secondary{background:var(--muted)}

                table.sample td[data-match="true"]{background:#90dba5!important}
                table.sample td[data-match="false"]{background:#e88a8a!important}

                @media print {.top-right-actions{display:none!important}}
            </style>
            </head><body>
            """
            )

            f.write("<div class='page-wrap'>")

            # Buttons
            f.write(
                """
            <div class="top-right-actions">
                <button id="download-pdf" class="btn-download">⬇️ Download PDF</button>
                <button class="btn-download btn-secondary" onclick="window.print()">🖨️ Print</button>
            </div>
            """
            )

            f.write("<div class='card' id='report-content'>")

            # Header
            f.write("<h1>Export Dashboard Validation Log</h1>")
            f.write("<div class='meta'>")
            f.write(
                f"<strong>Customer:</strong> {html.escape(selected_customer)} &nbsp; | &nbsp; "
            )
            f.write(
                f"<strong>Export:</strong> {html.escape(patient_id)} &nbsp; | &nbsp; "
            )
            f.write(
                f"<strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            )
            f.write("</div><hr>")

            # ================= FAILED CASES SECTION =================
            if failed_logs:
                f.write(
                    """
                <div style="margin:12px 0 18px 0;">
                    <h2 style="color:#b91c1c;">❌ Failed Cases Summary (click to jump)</h2>
                    <div style="
                        background:#fff5f5;
                        border:1px solid #f5c2c7;
                        border-radius:6px;
                        padding:10px;
                        max-height:240px;
                        overflow:auto;">
                """
                )

                for idx, e in enumerate(log_entries):
                    if "❌" in e:
                        f.write(
                            f"<div style='margin-bottom:6px;'>"
                            f"• <a href='#log-{idx}' style='color:#b91c1c;text-decoration:none;'>"
                            f"{html.escape(e)}</a></div>"
                        )

                f.write("</div></div>")
            else:
                f.write(
                    """
                <div style="margin:12px 0 18px 0;">
                    <h2 style="color:#15803d;">✅ No Failed Cases Detected</h2>
                </div>
                """
                )

            # ================= PASSED CASES SUMMARY =================
            if passed_logs:
                f.write(
                    """
                <div style="margin:12px 0 18px 0;">
                    <h2 style="color:#15803d;">✅ Passed Cases Summary</h2>
                    <div style="
                        background:#f0fdf4;
                        border:1px solid #bbf7d0;
                        border-radius:6px;
                        padding:10px;
                        max-height:200px;
                        overflow:auto;">
                """
                )

                for p in passed_logs:
                    f.write(
                        f"<div style='margin-bottom:6px;'>• {html.escape(p)}</div>"
                    )

                f.write("</div></div>")
            else:
                f.write(
                    """
                <div style="margin:12px 0 18px 0;">
                    <h2 style="color:#6b7280;">❌ No Passed Steps Recorded</h2>
                </div>
                """
                )

            # ================= FULL LOG ENTRIES =================
            for e in log_entries:
                f.write(f"<div class='entry'>{html.escape(e)}</div>")

            # ================= SAMPLE TABLE =================
            if sample_table_html:
                f.write(sample_table_html)

            f.write("</div></div>")

            # PDF JS
            f.write(
                """
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.9.3/html2pdf.bundle.min.js"></script>
            <script>
            document.getElementById('download-pdf').addEventListener('click', function(){
                var btn=this;
                btn.disabled=true;
                var el=document.getElementById('report-content');
                html2pdf().from(el).set({
                    margin:0.35,
                    filename:'export_report.pdf',
                    html2canvas:{scale:2},
                    jsPDF:{unit:'in',format:'a4',orientation:'portrait'}
                }).save().then(()=>btn.disabled=false)
                .catch(()=>{btn.disabled=false;window.print();});
            });
            </script>
            """
            )

            f.write("</body></html>")

        html_report_written = True
        log(f"✅ Log saved to {outpath.resolve()} (HTML with failed summary)")
    except Exception as ex:
        log(f"❌ Failed to save HTML log: {ex}")


# ─── Tkinter Progress UI ─────────────────────────────────────────────
class ProgressWindow:
    def __init__(self, master: Tk, total_steps: int) -> None:
        self.window = Toplevel(master)
        self.window.title("Export Dashboard Validation Progress")
        self.window.geometry("700x180")
        self.window.configure(bg="#4f8611")
        self.window.resizable(False, False)

        self.step = 0
        self.total_steps = max(1, total_steps)

        self.status_var = StringVar(value="Starting validation...")
        self.status_label = Label(
            self.window,
            textvariable=self.status_var,
            bg="#4f8611",
            fg="white",
            font=("Arial", 13, "bold"),
            wraplength=640,
            justify="center",
        )
        self.status_label.pack(pady=(12, 6), fill="x", padx=10)

        self.progress = tkttk.Progressbar(
            self.window, orient="horizontal", length=560, mode="determinate"
        )
        self.progress.pack(pady=(2, 6), padx=20, fill="x")
        self.progress["maximum"] = self.total_steps
        self.progress["value"] = 0

        self.percent_label = Label(
            self.window,
            text="0%",
            bg="#4f8611",
            fg="white",
            font=("Arial", 11, "bold"),
        )
        self.percent_label.pack(pady=(0, 6))

        self.last_msg = StringVar(value="")
        self.last_msg_label = Label(
            self.window,
            textvariable=self.last_msg,
            bg="#4f8611",
            fg="white",
            font=("Arial", 10),
            wraplength=640,
            justify="center",
        )
        self.last_msg_label.pack(pady=(0, 12), fill="x", padx=10)

        self.window.update()

    def update(self, step_description: str) -> None:
        self.step = min(self.step + 1, self.total_steps)
        self.status_var.set(step_description)

        percent = int((self.step / self.total_steps) * 100)
        self.progress["value"] = self.step
        self.percent_label.config(text=f"{percent}%")

        self.last_msg.set(step_description)
        log(step_description)
        self.window.update_idletasks()
        time.sleep(0.35)

    def complete(self) -> None:
        self.status_var.set("✅ Validation completed!")
        self.progress["value"] = self.total_steps
        self.percent_label.config(text="100%")
        log("Validation completed")
        self.window.update_idletasks()
        time.sleep(0.8)
        try:
            self.window.destroy()
        except Exception:
            pass


# ─── Core Selenium Classes ─────────────────────────────────────────────
class ConfParser:
    def __init__(self, config_file_path: Path) -> None:
        self.config_file_path = config_file_path
        self.config = ConfigParser()
        if not self.config_file_path.exists():
            raise FileNotFoundError(f"Config file not found: {self.config_file_path}")
        self.config.read(self.config_file_path)
        log(f"Config Parser read data from: {self.config_file_path}")


class ChromeDriverSetup(ConfParser):
    def __init__(self, config_file_path: Path):
        ConfParser.__init__(self, config_file_path)
        self.options = webdriver.ChromeOptions()

        try:
            self.options.add_argument(self.config["path"]["chrome_profile"])
        except Exception:
            pass

        prefs = {
            "download.default_directory": str(DOWNLOAD_DIR.resolve()),
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
        }
        self.options.add_experimental_option("prefs", prefs)

        chrome_driver_path = self.config["path"]["chrome_driver"]
        try:
            service = Service(executable_path=chrome_driver_path)
            self.driver = webdriver.Chrome(service=service, options=self.options)
        except Exception:
            self.driver = webdriver.Chrome(
                executable_path=chrome_driver_path, options=self.options
            )
        log("Chrome Driver Setup done.")


class SupportiveFunctions:
    driver: webdriver.Chrome  # type: ignore

    def ajax_preloader_wait(self, timeout: int = 300) -> None:
        try:
            time.sleep(0.8)
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element(
                    (By.XPATH, "//div[contains(@class,'ajax_preloader')]")
                )
            )
            time.sleep(0.4)
        except TimeoutException:
            log("Notice: ajax_preloader_wait timed out (element may persist).")
        except Exception as e:
            log(f"Notice: ajax_preloader_wait exception (may be safe): {e}")

    def get_url(self, section: str, key: str, fallback: str = "about:blank") -> str:
        try:
            return self.config.get(section, key)
        except Exception:
            log(f"❌ URL not found in config [{section}] → {key}. Using fallback.")
            return fallback

    def click_from_config(self, section: str, key: str):
        try:
            xpath = self.config.get(section, key)
            elem = WebDriverWait(self.driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            elem.click()
            return elem
        except Exception as e:
            log(f"❌ click_from_config failed [{section}] → {key}: {e}")
            raise


class CozevaLogin(ChromeDriverSetup, SupportiveFunctions):
    def certlogin_cozeva(
        self, selected_customer: str, progress: ProgressWindow
    ) -> None:
        try:
            progress.update("Logging into Cozeva (CERT)...")
            self.driver.get(self.config.get("cert", "logout_url", fallback="about:blank"))
            self.driver.get(self.config.get("cert", "login_url", fallback="about:blank"))
            self.driver.maximize_window()

            user = os.environ.get("CS2User")
            pwd = os.environ.get("CS2Password")
            if not all((user, pwd)):
                raise RuntimeError(
                    "Environment variables CS2User / CS2Password not set."
                )

            self.driver.find_element(By.ID, "edit-name").send_keys(user)
            self.driver.find_element(By.ID, "edit-pass").send_keys(pwd)
            self.driver.find_element(By.ID, "edit-submit").click()

            WebDriverWait(self.driver, 120).until(
                EC.presence_of_element_located((By.ID, "reason_textbox"))
            )
            self.driver.find_element(By.XPATH, "//*[@id='select-customer']").click()
            self.driver.find_element(
                By.XPATH, f"//*[contains(text(), '{str(selected_customer)}')]"
            ).click()

            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.ID, "reason_textbox"))
            )
            reason_text = self.config.get(
                "credentials", "export_reason", fallback=""
            )
            self.driver.find_element(By.ID, "reason_textbox").send_keys(reason_text)
            self.driver.find_element(By.ID, "edit-submit").click()
            time.sleep(5)

            self.ajax_preloader_wait()
            progress.update("Logged in successfully (CERT).")
        except Exception as e:
            log(f"❌ Login Error (CERT): {e}")
            messagebox.showerror("Login Error", str(e))
            raise

    def prodlogin_cozeva(
        self, selected_customer: str, progress: ProgressWindow
    ) -> None:
        try:
            progress.update("Logging into Cozeva (PROD)...")
            self.driver.get(self.config.get("prod", "logout_url", fallback="about:blank"))
            self.driver.get(self.config.get("prod", "login_url", fallback="about:blank"))
            self.driver.maximize_window()

            user = os.environ.get("CS2User")
            pwd = os.environ.get("CS2Password")
            if not all((user, pwd)):
                raise RuntimeError(
                    "Environment variables CS2User / CS2Password not set."
                )

            self.driver.find_element(By.ID, "edit-name").send_keys(user)
            self.driver.find_element(By.ID, "edit-pass").send_keys(pwd)
            self.driver.find_element(By.ID, "edit-submit").click()

            WebDriverWait(self.driver, 120).until(
                EC.presence_of_element_located((By.ID, "reason_textbox"))
            )
            self.driver.find_element(By.XPATH, "//*[@id='select-customer']").click()
            self.driver.find_element(
                By.XPATH, f"//*[contains(text(), '{str(selected_customer)}')]"
            ).click()

            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.ID, "reason_textbox"))
            )
            reason_text = self.config.get(
                "credentials", "export_reason", fallback=""
            )
            self.driver.find_element(By.ID, "reason_textbox").send_keys(reason_text)
            self.driver.find_element(By.ID, "edit-submit").click()
            time.sleep(5)

            self.ajax_preloader_wait()
            progress.update("Logged in successfully (PROD).")
        except Exception as e:
            log(f"❌ Login Error (PROD): {e}")
            messagebox.showerror("Login Error", str(e))
            raise

    def logout_cozeva(
        self,
        progress: ProgressWindow,
        selected_customer: Optional[str] = None,
        patient_id: Optional[str] = None,
    ) -> None:
        global html_report_written
        try:
            self.driver.get(self.config.get("cert", "logout_url", fallback="about:blank"))
            self.driver.quit()
            progress.complete()

            if not html_report_written:
                save_logs_to_html(
                    selected_customer or "Unknown", patient_id or "Unknown"
                )
            log("Logged out.")
        except Exception as e:
            log(f"❌ Logout Error: {e}")
            if not html_report_written:
                save_logs_to_html(
                    selected_customer or "Unknown", patient_id or "Unknown"
                )
            raise


class SticketCreation(CozevaLogin):
    def _click_sidenav_if_present(self) -> None:
        try:
            contact_log_tab = self.driver.find_element(
                By.XPATH, "//a[@id='contact_log_tab']"
            )
            if contact_log_tab.is_displayed() and contact_log_tab.is_enabled():
                log(
                    "contact_log_tab is present & clickable → skipping sidenav click."
                )
                return
            else:
                log(
                    "contact_log_tab found but NOT clickable → proceeding to sidenav click."
                )
        except NoSuchElementException:
            log("contact_log_tab not found → proceeding to sidenav click.")
        except Exception as e:
            log(f"Unexpected error while checking contact_log_tab: {e}")

        try:
            sidenav_btn = self.driver.find_element(
                By.XPATH, "//*[@data-target='sidenav_slide_out']"
            )
            if sidenav_btn.is_displayed() and sidenav_btn.is_enabled():
                sidenav_btn.click()
                log("Clicked sidenav_slide_out to open side navigation.")
                self.ajax_preloader_wait()
            else:
                log(
                    "sidenav_slide_out element found but NOT clickable; skipping click."
                )
        except NoSuchElementException:
            log("sidenav_slide_out not present; assuming side nav already open.")
        except ElementNotInteractableException as e:
            log(f"sidenav_slide_out present but not interactable: {e}")
        except Exception as e:
            log(
                f"Warning: unexpected error while trying to click sidenav_slide_out: {e}"
            )

    def sticket_creation(self) -> None:
        try:
            log("▶️ Opening Patient Dashboard page")
            self.driver.get(self.get_url("patient_dashboard", "patient", fallback="about:blank"))
            self.ajax_preloader_wait()
            log("✅ Patient Dashboard loaded")

            log("▶️ Clicking patient dropdown")
            self.click_from_config("PatientDashboard", "xpath_patient_dropdown")
            time.sleep(2)
            log("✅ Patient dropdown opened")

            actions = ActionChains(self.driver)
            messages_xpath = self.config.get("PatientDashboard", "xpath_messages")
            new_sticket_xpath = self.config.get("PatientDashboard", "xpath_new_sticket")

            log("▶️ Hovering over Messages menu")
            messages_elem = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, messages_xpath))
            )
            actions.move_to_element(messages_elem).perform()
            time.sleep(2)
            log("✅ Messages menu visible")

            log("▶️ Clicking New Sticket")
            new_sticket_elem = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, new_sticket_xpath))
            )
            actions.move_to_element(new_sticket_elem).pause(0.5).click().perform()
            self.ajax_preloader_wait()
            log("✅ New sticket popup opened")

            created_sticket_text = "New sticket for testing purpose CozevaQA"
            log("▶️ Entering sticket text")
            sticket_textbox = self.click_from_config("PatientDashboard", "xpath_sticket_text_box")
            sticket_textbox.clear()
            sticket_textbox.send_keys(created_sticket_text)
            log(f"✅ Sticket text entered: {created_sticket_text}")

            log("▶️ Saving sticket")
            self.click_from_config("PatientDashboard", "xpath_sticket_save")
            log("✅ Sticket saved")

            log("▶️ Closing sticket dialog")
            close_xpath = self.config.get("PatientDashboard", "xpath_sticket_close")
            close_btn = WebDriverWait(self.driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, close_xpath))
            )
            close_btn.click()
            time.sleep(2)
            log("✅ Sticket dialog closed")

            # Navigate to sticket log
            self.click_from_config("LOCATOR", "xpath_sidemenu_click")
            time.sleep(2)
            self.click_from_config("SideBar", "xpath_sticket_log")
            self.ajax_preloader_wait()
            time.sleep(3)

            latest_note_xpath = self.config.get("SticketLog", "xpath_latest_note")
            notes_elements = WebDriverWait(self.driver, 15).until(
                EC.visibility_of_all_elements_located((By.XPATH, latest_note_xpath))
            )
            log(f"Total sticket notes found: {len(notes_elements)}")

            matched_row = None
            for note in notes_elements:
                note_text = note.text.strip()
                log(f"Sticket Note: {note_text}")
                if created_sticket_text.lower() in note_text.lower():
                    log("✅ Matching sticket row found")
                    matched_row = note.find_element(By.XPATH, "./ancestor::tr")
                    break

            if not matched_row:
                raise AssertionError("❌ Matching sticket row not found")

            # Click patient link from matched row
            patient_link_xpath = self.config.get("SticketLog", "xpath_patient_link")
            if patient_link_xpath.startswith("//"):
                patient_link_xpath = "." + patient_link_xpath

            patient_link = matched_row.find_element(By.XPATH, patient_link_xpath)
            patient_name = patient_link.text.strip()
            log(f"✅ Patient link found: {patient_name}")

            parent_window = self.driver.current_window_handle
            existing_windows = self.driver.window_handles.copy()

            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", patient_link
            )
            time.sleep(1)
            patient_link.click()

            WebDriverWait(self.driver, 15).until(
                lambda d: len(d.window_handles) > len(existing_windows)
            )

            new_window = [w for w in self.driver.window_handles if w not in existing_windows][0]
            self.driver.switch_to.window(new_window)
            log("✅ Switched to patient window")

            self.ajax_preloader_wait()

            # Edit sticket
            self.click_from_config("PatientDashboard", "xpath_dashboard_sticket_icon")
            self.ajax_preloader_wait()

            self.click_from_config("PatientDashboard", "xpath_sticket_kebab")
            time.sleep(1)
            self.click_from_config("PatientDashboard", "xpath_sticket_edit")
            time.sleep(2)

            edit_suffix = " editing sticket for testing purpose"
            edit_box = self.click_from_config("PatientDashboard", "xpath_sticket_text_box")

            existing_text = edit_box.get_attribute("value") or ""
            edit_box.clear()
            updated_text = existing_text + edit_suffix
            edit_box.send_keys(updated_text)

            self.click_from_config("PatientDashboard", "xpath_sticket_save")
            time.sleep(10)

            edited_text_elem = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located(
                    (By.XPATH, self.config.get("PatientDashboard", "xpath_sticket_text"))
                )
            )
            edited_text = edited_text_elem.text.strip()
            log(f"✅ Edited sticket text: {edited_text}")

            time.sleep(10)

            # Delete sticket
            self.click_from_config("PatientDashboard", "xpath_sticket_kebab")
            time.sleep(1)
            self.click_from_config("PatientDashboard", "xpath_sticket_delete")
            time.sleep(1)
            self.click_from_config("PatientDashboard", "xpath_delete_yes")
            self.ajax_preloader_wait()
            log("✅ Sticket deleted successfully")

            # Close patient window and return
            self.driver.close()
            self.driver.switch_to.window(parent_window)
            log("✅ Switched back to main window")

            # Verify deletion in sticket log
            self.driver.refresh()
            self.ajax_preloader_wait()
            time.sleep(3)

            refreshed_notes = WebDriverWait(self.driver, 15).until(
                EC.visibility_of_all_elements_located((By.XPATH, latest_note_xpath))
            )

            found_after_delete = False
            for note in refreshed_notes:
                if edited_text.lower() in note.text.lower():
                    found_after_delete = True
                    break

            if found_after_delete:
                raise AssertionError("❌ Edited sticket still exists after deletion")
            else:
                log("✅ Edited sticket NOT found after deletion — Test Passed")

        except Exception as e:
            log(f"❌ Sticket validation failed: {e}")
            raise


def run_sticket_creation(
    selected_customer: str, patient_id: str, selected_env: str, root: Tk
) -> None:
    progress = ProgressWindow(root, total_steps=5)
    runner: SticketCreation | None = None
    try:
        progress.update("Initializing browser...")
        runner = SticketCreation(CONFIG_FILE_PATH)

        if selected_env.upper() == "CERT":
            runner.certlogin_cozeva(selected_customer, progress)
        else:
            runner.prodlogin_cozeva(selected_customer, progress)

        progress.update("Creating sticket...")
        runner.sticket_creation()

        progress.update("Sticket flow completed successfully ✅")

    except Exception as e:
        log(f"❌ Execution failed: {e}")
        messagebox.showerror("Execution Error", str(e))

    finally:
        if runner:
            try:
                runner.logout_cozeva(progress, selected_customer, patient_id)
            except Exception:
                pass


def main() -> None:
    root = Tk()
    root.withdraw()

    from Export_DashboardUI import sticket_ui

    selected_customer, patient_id, selected_env = sticket_ui(root)
    if not selected_customer or not patient_id or not selected_env:
        log("No selection made. Exiting.")
        root.destroy()
        return

    run_sticket_creation(selected_customer, patient_id, selected_env, root)

    try:
        root.destroy()
    except Exception:
        pass

