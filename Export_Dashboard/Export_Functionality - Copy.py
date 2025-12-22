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
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from tkinter import Tk, Toplevel, Label, StringVar, messagebox
from tkinter import ttk as tkttk

# ─── Configuration ─────────────────────────────────────────────
CONFIG_FILE_PATH = Path(r"C:\Users\nsikder\Downloads\config.ini")
DOWNLOAD_DIR = Path(r"C:\Users\nsikder\PycharmProjects\Export Dashboard\Exported Files")
LOG_HTML_FILE = Path("validation_log.html")

# ─── Global log store ─────────────────────────────────────────────
log_entries: list[str] = []
html_report_written: bool = False  # avoid overwriting report once written


def log(message: str) -> None:
    """Append message to global log (timestamped) and also print to stdout."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{ts}] {message}"
    log_entries.append(entry)
    print(entry)


def save_logs_to_html(customer: str,
                      export_type: str,
                      filename: Path = LOG_HTML_FILE,
                      sample_table_html: Optional[str] = None) -> None:
    """
    Write log_entries to a styled HTML file and optionally insert a sample_table_html fragment.
    Adds a 'Download PDF' button in the top-right corner of the report card which uses
    html2pdf (html2pdf.bundle.min.js) to convert the report to a paginated PDF and trigger a download.
    Falls back to window.print() if the library can't load.
    """
    global html_report_written
    try:
        outpath = Path(filename)
        with outpath.open("w", encoding="utf-8") as f:
            f.write("<!doctype html><html><head><meta charset='utf-8'><title>Export Validation Log</title>")
            f.write("""<style>
                :root{--brand:#2f6f17;--muted:#4a5568;--bg:#f6f7f9}
                body{font-family:Segoe UI, Arial, sans-serif;background:var(--bg);padding:20px;margin:0}
                .page-wrap{max-width:1100px;margin:20px auto;position:relative}
                .card{background:#fff;border-radius:8px;padding:22px 22px 30px 22px;box-shadow:0 2px 6px rgba(0,0,0,0.08);position:relative;overflow:visible}
                h1{color:var(--brand);margin:0 0 10px 0}
                .meta{color:#555;font-size:0.95rem;margin-bottom:12px}
                .entry{font-family:Times New Roman;background:#f3f4f6;padding:8px;margin:6px 0;border-radius:6px;overflow:auto}
                .success{color:green}.error{color:red}
                table.sample{border-collapse:collapse;width:100%;font-family:Segoe UI, Arial, sans-serif;margin-top:12px}
                table.sample th, table.sample td{border:1px solid #e9ecef;padding:6px;vertical-align:top;text-align:left}
                table.sample thead th{background:#f3f4f6;font-weight:600}

                /* top-right button group */
                .top-right-actions{position:absolute;top:12px;right:12px;display:flex;gap:8px;align-items:center;z-index:10}
                .btn-download{display:inline-block;padding:8px 12px;border-radius:6px;background:var(--brand);color:white;text-decoration:none;border:none;cursor:pointer;font-weight:600;box-shadow:0 1px 3px rgba(0,0,0,0.12)}
                .btn-secondary{background:var(--muted)}
                .pdf-note{font-size:0.85rem;color:#333;margin-right:auto;display:none} /* hidden in top-right layout */
                @media print {
                  .top-right-actions{display:none !important}
                }
                </style></head><body>""")

            # page wrapper (positions top-right actions relative to this)
            f.write("<div class='page-wrap'>")

            # top-right actions: placed before card so it's still inside .page-wrap
            f.write("""
                <div class="top-right-actions" aria-hidden="false">
                  <button id="download-pdf" class="btn-download" title="Download PDF">⬇️ Download PDF</button>
                  <button id="print-pdf" class="btn-download btn-secondary" title="Print" onclick="window.print();">🖨️ Print</button>
                </div>
            """)

            # A wrapper that will be converted to PDF (the card)
            f.write("<div class='card' id='report-content'>")
            f.write("<h1>Export Dashboard Validation Log</h1>")
            f.write("<div class='meta'>")
            f.write(f"<strong>Customer:</strong> {html.escape(str(customer))} &nbsp; | &nbsp; ")
            f.write(f"<strong>Export:</strong> {html.escape(str(export_type))} &nbsp; | &nbsp; ")
            f.write(f"<strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            f.write("</div><hr>")

            # write log entries
            for e in log_entries:
                f.write(f"<div class='entry'>{html.escape(e)}</div>\n")

            # insert the sample table HTML if provided
            if sample_table_html:
                f.write(sample_table_html)

            f.write("</div>")  # end .card
            f.write("</div>")  # end .page-wrap

            # Add html2pdf script and inline JS to wire the top-right download button
            f.write("""
            <!-- html2pdf (bundles html2canvas + jsPDF) from CDN -->
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.9.3/html2pdf.bundle.min.js"></script>
            <script>
            (function(){
                var downloadBtn = document.getElementById('download-pdf');

                function fallback_print(){
                    alert('Automatic PDF generation failed or offline. The browser print dialog will open; choose "Save as PDF" to save.');
                    window.print();
                }

                downloadBtn.addEventListener('click', function(){
                    downloadBtn.disabled = true;
                    var originalText = downloadBtn.innerText;
                    downloadBtn.innerText = 'Generating PDF...';

                    var element = document.getElementById('report-content');

                    if (typeof html2pdf === 'undefined') {
                        console.warn('html2pdf not available; falling back to print.');
                        downloadBtn.innerText = originalText;
                        downloadBtn.disabled = false;
                        fallback_print();
                        return;
                    }

                    var opt = {
                        margin:       0.35,
                        filename:     'export_report.pdf',
                        image:        { type: 'jpeg', quality: 0.98 },
                        html2canvas:  { scale: 2, useCORS: true, allowTaint: true },
                        jsPDF:        { unit: 'in', format: 'a4', orientation: 'portrait' }
                    };

                    html2pdf().set(opt).from(element).save().then(function(){
                        downloadBtn.innerText = originalText;
                        downloadBtn.disabled = false;
                    }).catch(function(err){
                        console.error('html2pdf error:', err);
                        downloadBtn.innerText = originalText;
                        downloadBtn.disabled = false;
                        fallback_print();
                    });
                });
            })();
            </script>
            """)

            f.write("</body></html>")

        html_report_written = True
        log(f"✅ Log saved to {outpath.resolve()} (HTML with top-right PDF/Print buttons)")
    except Exception as ex:
        log(f"❌ Failed to save HTML log: {ex}")


# ─── Tkinter Progress UI ─────────────────────────────────────────────
class ProgressWindow:
    def __init__(self, master: Tk, total_steps: int) -> None:
        # slightly wider/taller window to reduce wrap collisions
        self.window = Toplevel(master)
        self.window.title("Export Dashboard Validation Progress")
        self.window.geometry("700x180")
        self.window.configure(bg="#4f8611")
        self.window.resizable(False, False)

        self.step = 0
        self.total_steps = max(1, total_steps)

        self.status_var = StringVar(value="Starting validation...")
        # status label with wraplength and centered justification so long text wraps
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

        # progressbar with horizontal padding so labels don't touch it
        self.progress = tkttk.Progressbar(self.window, orient="horizontal", length=560, mode="determinate")
        self.progress.pack(pady=(2, 6), padx=20, fill="x")
        self.progress["maximum"] = self.total_steps
        self.progress["value"] = 0

        # percent label placed below the progress bar to avoid overlap
        self.percent_label = Label(
            self.window,
            text="0%",
            bg="#4f8611",
            fg="white",
            font=("Arial", 11, "bold")
        )
        self.percent_label.pack(pady=(0, 6))

        # Properly keep a reference to the label displaying the last message.
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
        # add a little bottom padding to prevent crowding
        self.last_msg_label.pack(pady=(0, 12), fill="x", padx=10)

        self.window.update()

    def update(self, step_description: str) -> None:
        """Update progress bar and UI + append to logs."""
        # advance step (if you prefer manual control later, we can add an increment param)
        self.step = min(self.step + 1, self.total_steps)
        self.status_var.set(step_description)

        percent = int((self.step / self.total_steps) * 100)
        self.progress["value"] = self.step
        self.percent_label.config(text=f"{percent}%")

        # update the persistent last_msg_label rather than destroying widgets
        self.last_msg.set(step_description)
        log(step_description)
        self.window.update_idletasks()
        time.sleep(0.35)

    def complete(self) -> None:
        """Mark complete, log, wait briefly and destroy window."""
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

        # chrome_profile in config should typically contain something like --user-data-dir=...
        try:
            self.options.add_argument(self.config['path']['chrome_profile'])
        except Exception:
            pass

        prefs = {
            "download.default_directory": str(DOWNLOAD_DIR.resolve()),
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True,
        }
        self.options.add_experimental_option("prefs", prefs)

        chrome_driver_path = self.config['path']['chrome_driver']
        try:
            service = Service(executable_path=chrome_driver_path)
            self.driver = webdriver.Chrome(service=service, options=self.options)
        except Exception:
            # fallback older style if needed
            self.driver = webdriver.Chrome(executable_path=chrome_driver_path, options=self.options)
        log("Chrome Driver Setup done.")


class SupportiveFunctions:
    driver: webdriver.Chrome  # type: ignore

    def ajax_preloader_wait(self, timeout: int = 300) -> None:
        """Wait for ajax preloader to disappear; tolerant to minor failures."""
        try:
            time.sleep(0.8)
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element((By.XPATH, "//div[contains(@class,'ajax_preloader')]"))
            )
            time.sleep(0.4)
        except TimeoutException:
            log("Notice: ajax_preloader_wait timed out (element may persist).")
        except Exception as e:
            log(f"Notice: ajax_preloader_wait exception (may be safe): {e}")


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
            reason_text = self.config.get("credentials", "export_reason", fallback="")
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
            reason_text = self.config.get("credentials", "export_reason", fallback="")
            self.driver.find_element(By.ID, "reason_textbox").send_keys(reason_text)
            self.driver.find_element(By.ID, "edit-submit").click()
            time.sleep(5)

            self.ajax_preloader_wait()
            progress.update("Logged in successfully (PROD).")
        except Exception as e:
            log(f"❌ Login Error (PROD): {e}")
            messagebox.showerror("Login Error", str(e))
            raise

    def logout_cozeva(self, progress: ProgressWindow,
                      customer: Optional[str] = None,
                      export_type: Optional[str] = None) -> None:
        """Logout and close driver; ensure logs are saved (but don't overwrite report if already written)."""
        global html_report_written
        try:
            # Using cert logout_url here as before; adjust if needed per env.
            self.driver.get(self.config.get("cert", "logout_url", fallback="about:blank"))
            self.driver.quit()
            progress.complete()

            # Only save HTML if not already written (e.g., from export_dashboard with table)
            if not html_report_written:
                save_logs_to_html(customer or "Unknown", export_type or "Unknown")
            log("Logged out.")
        except Exception as e:
            log(f"❌ Logout Error: {e}")
            if not html_report_written:
                save_logs_to_html(customer or "Unknown", export_type or "Unknown")
            raise


class ContactExport(CozevaLogin):
    def _click_sidenav_if_present(self) -> None:
        """
        Click the side-nav toggle only if the element //*[@data-target='sidenav_slide_out']
        is present and displayed. If it's not present, we assume the side nav is already open.
        """
        try:
            sidenav_btn = self.driver.find_element(By.XPATH, "//*[@data-target='sidenav_slide_out']")
            if sidenav_btn.is_displayed() and sidenav_btn.is_enabled():
                sidenav_btn.click()
                log("Clicked sidenav_slide_out to open side navigation.")
                self.ajax_preloader_wait()
            else:
                log("sidenav_slide_out element found but not clickable; skipping click.")
        except NoSuchElementException:
            log("sidenav_slide_out not present; assuming side nav already open.")
        except ElementNotInteractableException as e:
            log(f"sidenav_slide_out present but not interactable: {e}")
        except Exception as e:
            log(f"Warning: unexpected error while trying to click sidenav_slide_out: {e}")

    def _capture_ui_rows_for_headers(self, header_names: list[str], max_rows: int = 10) -> list[list[str]]:
        """
        On the already-open sticket log page in the main window, capture up to `max_rows`
        rows of data for the given header_names (by matching header text).

        Assumes you have already switched to the correct window and refreshed the page.
        Returns a list of rows, each row being list[str] aligned with header_names.
        """
        def _norm(s: str) -> str:
            return (s or "").strip().lower()

        # Find a table whose headers overlap with our header_names
        tables = self.driver.find_elements(By.XPATH, "//table")
        target_table = None
        target_header_texts: list[str] = []

        for tbl in tables:
            try:
                ths = tbl.find_elements(By.XPATH, ".//th")
                if not ths:
                    continue
                th_texts_raw = [th.text for th in ths]
                th_texts_norm = [_norm(t) for t in th_texts_raw]

                # count how many of our header_names appear in this table's headers
                overlap = 0
                for hn in header_names:
                    hn_norm = _norm(hn)
                    if any(hn_norm in t or t in hn_norm for t in th_texts_norm):
                        overlap += 1

                if overlap >= max(1, len(header_names) // 2):
                    target_table = tbl
                    target_header_texts = th_texts_raw
                    break
            except Exception:
                continue

        if target_table is None:
            raise RuntimeError("Could not locate sticket log table for UI comparison.")

        log("Found candidate sticket log table for comparison.")

        # Build a mapping from header text -> index for the UI table
        ui_header_index_map: dict[str, int] = {}
        norm_header_texts = [_norm(t) for t in target_header_texts]
        for idx, t in enumerate(norm_header_texts):
            ui_header_index_map[t] = idx

        ui_indices: list[int] = []
        for hn in header_names:
            hn_norm = _norm(hn)
            idx = None
            if hn_norm in ui_header_index_map:
                idx = ui_header_index_map[hn_norm]
            else:
                # fuzzy/substring match
                for i, t in enumerate(norm_header_texts):
                    if hn_norm in t or t in hn_norm:
                        idx = i
                        break

            if idx is None:
                ui_indices.append(-1)
                log(f"Notice: header '{hn}' not found in sticket UI table; will leave blank for comparison.")
            else:
                ui_indices.append(idx)

        # Capture up to max_rows rows from tbody
        rows_ui: list[list[str]] = []
        trs = target_table.find_elements(By.XPATH, ".//tbody/tr")
        for tr in trs[:max_rows]:
            tds = tr.find_elements(By.XPATH, "./td")
            row_vals: list[str] = []
            for idx in ui_indices:
                if 0 <= idx < len(tds):
                    row_vals.append((tds[idx].text or "").strip())
                else:
                    row_vals.append("")
            rows_ui.append(row_vals)

        log(f"Captured {len(rows_ui)} rows from sticket UI for comparison.")
        return rows_ui

    def contact_export(self, progress: ProgressWindow) -> None:
        progress.update("Running Contact Export...")
        DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

        # Click sidenav only if the toggle is present
        self._click_sidenav_if_present()

        # Make sure any loaders are gone
        self.ajax_preloader_wait()

        # Now safely wait until the tab is actually clickable
        try:
            contact_tab = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@id='contact_log_tab']"))
            )
            contact_tab.click()
        except Exception as e:
            log(f"❌ Could not click contact_log_tab: {e}")
            raise

        self.ajax_preloader_wait()
        self.driver.find_element(By.XPATH, "//*[@data-target='datatable_bulk_filter_0_contact_log']").click()
        self.driver.find_element(By.XPATH, "//a[contains(text(), 'Export all to CSV')]").click()
        self.driver.find_element(By.XPATH, "//a[normalize-space(text())='YES']").click()
        progress.update("Contact export triggered.")

    def sticket_export(self, progress: ProgressWindow) -> None:
        progress.update("Running Sticket Export...")
        DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

        # Click sidenav only if the toggle is present
        self._click_sidenav_if_present()

        self.ajax_preloader_wait()

        try:
            sticket_tab = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@id='sticket_log_tab']"))
            )
            sticket_tab.click()
        except Exception as e:
            log(f"❌ Could not click sticket_log_tab: {e}")
            raise

        self.ajax_preloader_wait()
        self.driver.find_element(By.XPATH, "//*[@data-target='datatable_bulk_filter_0_sticket_log']").click()
        self.driver.find_element(By.XPATH, "//a[contains(text(), 'Export all to CSV')]").click()
        self.driver.find_element(By.XPATH, "//a[normalize-space(text())='YES']").click()
        progress.update("Sticket export triggered.")

    def export_dashboard(self, selected_customer: str, selected_export: str, progress: ProgressWindow) -> None:
        """
        Open export dashboard (in new window), poll status until completion, download and validate CSV,
        capture up to 10 rows of specified columns (selected by header name), exclude certain columns by header name,
        and insert an HTML table into the HTML log. Row data is NOT logged to console; it only appears in the HTML table.

        For Sticket exports, compares those rows against the already-open Sticket log UI page
        and colors cells green (match) or red (mismatch) in the HTML.
        """
        global html_report_written
        progress.update("Opening Export Dashboard...")
        original_windows = []
        original_window = None

        try:
            # open side nav
            self.driver.find_element(By.XPATH, "//*[@data-target='sidenav_slide_out']").click()
            time.sleep(0.8)

            original_windows = list(self.driver.window_handles)
            original_window = self.driver.current_window_handle

            # click data_validate (may open new window/tab)
            try:
                self.driver.find_element(By.XPATH, "//a[@id='data_validate']").click()
            except Exception as e_click:
                log(f"Warning: normal click for data_validate failed: {e_click}; trying JS click")
                try:
                    el = self.driver.find_element(By.XPATH, "//a[@id='data_validate']")
                    self.driver.execute_script("arguments[0].click();", el)
                except Exception as e_js:
                    log(f"❌ Could not click data_validate link: {e_js}")
                    raise

            # wait for new window and switch
            try:
                WebDriverWait(self.driver, 20).until(EC.new_window_is_opened(original_windows))
            except Exception:
                log("Notice: no new window detected after clicking data_validate — continuing in current window.")
            else:
                new_handles = [h for h in self.driver.window_handles if h not in original_windows]
                if new_handles:
                    new_window = new_handles[-1]
                    log(f"Switching to new window: {new_window}")
                    self.driver.switch_to.window(new_window)
                else:
                    last_handle = self.driver.window_handles[-1]
                    log(f"No diff handles found; switching to last handle: {last_handle}")
                    self.driver.switch_to.window(last_handle)

            self.ajax_preloader_wait()
            progress.update("Validating Export Dashboard data...")

            # read first table values defensively (customer column)
            cell = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//tr[@role='row' and contains(@class,'odd')][1]/td[4]"))
            )
            value = (cell.text or "").strip()
            progress.update("Validating Export Customer name...")
            if value.lower() == selected_customer.strip().lower():
                log(f"✅ Match found! Customer '{value}' matches selected '{selected_customer}'.")
            else:
                log(f"❌ Mismatch! Table shows '{value}' but selected '{selected_customer}'.")

            # read export type text from dashboard (may differ from selected_export)
            export_text_el = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH,
                                                "(//td[contains(@class, 'export-dashboard-row') "
                                                "and contains(@class, 'export-dashboard-row_pt')])[3]"))
            )
            export_type = (export_text_el.text or "").strip()
            log("Export type is: " + export_type)

            # poll status until percent == 100 or terminal state
            download_flag = False
            while not download_flag:
                status_el = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@class='status-info']"))
                )
                status_value = status_el.text or ""
                log("Status values (raw): " + status_value.replace("\n", " | "))
                status_values = status_value.split("\n", maxsplit=2)

                if len(status_values) < 3:
                    log("❌ Unexpected status text format, retrying after wait...")
                    time.sleep(4)
                    self.driver.refresh()
                    self.ajax_preloader_wait()
                    continue

                status_str = status_values[1].strip()
                percent_str = status_values[2].strip().strip("%")
                try:
                    percent = int(percent_str)
                except Exception:
                    percent = 0

                log(f"Current status: '{status_str}', Percent: {percent}%")
                progress.update("Validating Export dashboard percentage status...")

                if status_str in ("Deleted", "Unsuccessful"):
                    log(f"❌ Export ended in terminal state: {status_str}")
                    raise Exception(f"Export ended in terminal state: {status_str}")

                if percent < 100:
                    log(f"Progress {percent}% - waiting and refreshing...")
                    time.sleep(6)
                    self.driver.refresh()
                    self.ajax_preloader_wait()
                    continue

                # percent == 100
                download_flag = True
                if status_str == "Success":
                    log("✅ Export reported success; attempting download...")

                    # try to click download link (several fallbacks)
                    clicked = False
                    try:
                        dl = self.driver.find_element(By.XPATH, "(//a[contains(@href, 'unified_file_download')])[1]")
                        try:
                            dl.click()
                            clicked = True
                            log("Clicked download link (normal click).")
                        except Exception as e_click_dl:
                            log(f"Normal click on download link failed: {e_click_dl} - trying JS click")
                            try:
                                self.driver.execute_script("arguments[0].click();", dl)
                                clicked = True
                                log("Clicked download link via JS.")
                            except Exception as e_js_dl:
                                log(f"JS click also failed for download link: {e_js_dl}")
                    except Exception as e_find_dl:
                        log(f"❌ Could not find download link element: {e_find_dl}")

                    if not clicked:
                        log("❌ Could not click download link (all fallbacks). Continuing to poll for file but this may fail.")

                    # wait briefly for download to start
                    time.sleep(6)

                    # detect latest CSV (ignore .crdownload). extend timeout if needed
                    timeout_seconds = 60
                    file_path: Optional[Path] = None
                    for _ in range(timeout_seconds):
                        try:
                            files = list(DOWNLOAD_DIR.glob("*.csv"))
                            if files:
                                candidate = max(files, key=lambda p: p.stat().st_ctime)
                                crdownload_path = candidate.with_suffix(candidate.suffix + ".crdownload")
                                if not crdownload_path.exists():
                                    file_path = candidate
                                    break
                        except Exception as e_glob:
                            log(f"Warning while scanning downloads: {e_glob}")
                        time.sleep(1)

                    if not file_path:
                        log("❌ CSV file not found or download incomplete.")
                        raise Exception("CSV file not found or download incomplete.")

                    log(f"✅ CSV downloaded: {file_path}")

                    # ------------------ READ CSV robustly and select columns by header name ------------------
                    progress.update("Validating Exported file and columns...")
                    rows_sample: list[list[str]] = []
                    header_names: list[str] = []

                    # headers to always exclude (case-insensitive)
                    EXCLUDE_HEADERS = {
                        "patient", "dob", "member id", "member phone #", "member uid", "searchable member id",
                        "member fname", "member lname", "gender"
                    }

                    # --------- DEFINE which headers to include per export type ----------
                    INCLUDE_HEADERS_BY_EXPORT = {
                        "contact": [
                            "Member CozevaID", "Measure Details", "Encounter Datetime", "Route",
                            "Encounter Details", "Encounter Note", "With Whom", "Submitter", "PCP",
                            "Practice", "Health Plan", "Campaign", "Data Source"
                        ],
                        "sticket": [
                            "Created", "Last Updated", "Created by", "Last Updated by",
                            "PCP", "Latest Note", "Health Plan"
                        ],
                    }
                    # -------------------------------------------------------------------

                    with file_path.open("r", encoding="utf-8", newline="") as f:
                        sample = f.read(8192)
                        f.seek(0)

                        # sniff dialect
                        try:
                            dialect = csv.Sniffer().sniff(sample)
                            has_header = csv.Sniffer().has_header(sample)
                        except Exception:
                            dialect = csv.excel
                            has_header = True

                        detected_delim = getattr(dialect, "delimiter", ",")

                        reader = csv.reader(f, dialect)

                        # parse header row
                        headers = next(reader, None)
                        raw_headers = headers
                        if isinstance(raw_headers, list) and len(raw_headers) == 1 and isinstance(raw_headers[0], str):
                            single = raw_headers[0]
                            if detected_delim and detected_delim in single:
                                raw_headers = [h.strip() for h in single.split(detected_delim)]
                            else:
                                for d in [',', '|', ';', '\t']:
                                    if d in single:
                                        raw_headers = [h.strip() for h in single.split(d)]
                                        detected_delim = d
                                        log(f"Fallback-split header using delimiter {repr(d)}")
                                        break

                        # strip BOM
                        if raw_headers and isinstance(raw_headers, list) and len(raw_headers) > 0 and isinstance(
                                raw_headers[0], str) and raw_headers[0].startswith("\ufeff"):
                            raw_headers[0] = raw_headers[0].lstrip("\ufeff")

                        # Prepare normalized header list and map
                        def _norm(h: str) -> str:
                            return (h or "").strip().lower()

                        raw_lower = [_norm(h) for h in (raw_headers or [])]
                        header_index_map = {h: i for i, h in enumerate(raw_lower)}

                        # Decide which include-list to use based on SELECTED export in UI
                        export_kind_norm = _norm(selected_export)

                        chosen_include_list = None
                        for key, hdr_list in INCLUDE_HEADERS_BY_EXPORT.items():
                            if key in export_kind_norm:
                                chosen_include_list = hdr_list
                                break

                        if not chosen_include_list:
                            # No explicit includes provided for this export — capture all headers except excluded ones
                            log(f"Notice: no include-list found for '{selected_export}'. Capturing all non-excluded headers.")
                            selected_indices = [i for i, h in enumerate(raw_lower) if h not in EXCLUDE_HEADERS]
                        else:
                            # Build selected indices from chosen header names (case-insensitive)
                            selected_indices = []
                            for want_name in chosen_include_list:
                                want_norm = _norm(want_name)
                                if want_norm in header_index_map:
                                    selected_indices.append(header_index_map[want_norm])
                                else:
                                    # fuzzy/substring match as a best-effort
                                    matched = False
                                    for i, h in enumerate(raw_lower):
                                        if want_norm in h or h in want_norm:
                                            selected_indices.append(i)
                                            matched = True
                                            break
                                    if not matched:
                                        log(f"Notice: desired header '{want_name}' not found in CSV headers.")

                            # deduplicate while preserving order
                            seen = set()
                            selected_indices = [x for x in selected_indices if not (x in seen or seen.add(x))]

                        # Filter out any indices that map to excluded headers
                        filtered_indices = [idx for idx in selected_indices if
                                            idx < len(raw_lower) and raw_lower[idx] not in EXCLUDE_HEADERS]

                        # Build final header names for HTML (keep original header text)
                        header_names = [
                            raw_headers[i] if raw_headers and i < len(raw_headers) else f"Col {i}"
                            for i in filtered_indices
                        ]

                        # helper to detect repeated header rows on selected columns
                        def is_header_row(check_row_list: list[str]) -> bool:
                            if not raw_headers or not isinstance(check_row_list, list):
                                return False
                            left = [(raw_headers[i].strip().lower() if i < len(raw_headers) else "").strip()
                                    for i in filtered_indices]
                            right = [(check_row_list[i].strip().lower() if i < len(check_row_list) else "").strip()
                                     for i in filtered_indices]
                            return left == right

                        # collect up to 10 data rows (excluding repeated headers)
                        collected = 0
                        for row in reader:
                            # try to split malformed single-field rows
                            if isinstance(row, list) and len(row) == 1 and isinstance(row[0], str):
                                single = row[0]
                                if detected_delim and detected_delim in single:
                                    row = [c.strip() for c in single.split(detected_delim)]
                                else:
                                    splitted = None
                                    for d in [',', '|', ';', '\t']:
                                        if d in single:
                                            splitted = [c.strip() for c in single.split(d)]
                                            detected_delim = d
                                            log(f"Fallback-split data row using delimiter {repr(d)}")
                                            break
                                    if splitted is not None:
                                        row = splitted
                                    else:
                                        row = [single]

                            # skip repeated header rows that match on filtered_indices
                            try:
                                if is_header_row(row):
                                    continue
                            except Exception:
                                pass

                            # collect values for filtered_indices (but DO NOT log them to console)
                            vals = [row[i] if i < len(row) else "" for i in filtered_indices]

                            # add to rows_sample (only used for HTML table)
                            rows_sample.append(vals)
                            collected += 1
                            if collected >= 10:
                                break

                    log(f"Captured {len(rows_sample)} sample rows from CSV (filtered columns).")

                    # ---------- COMPARE AGAINST STICKET LOG UI (only for sticket exports) ----------
                    ui_match_matrix = None  # list[list[bool]] same shape as rows_sample
                    try:
                        # Only do comparison for Sticket exports (based on UI selected export)
                        export_kind_norm = (selected_export or "").strip().lower()
                        if "sticket" in export_kind_norm:
                            if original_window and original_window in self.driver.window_handles:
                                self.driver.switch_to.window(original_window)
                                log(f"Switched back to original window {original_window} for UI comparison.")
                                # Just refresh the already-open sticket log page
                                self.driver.refresh()
                                self.ajax_preloader_wait()
                            else:
                                log("Original window handle not available for UI comparison.")

                            ui_rows = self._capture_ui_rows_for_headers(
                                header_names, max_rows=len(rows_sample)
                            )

                            ui_match_matrix = []
                            for i, csv_row in enumerate(rows_sample):
                                ui_row = ui_rows[i] if i < len(ui_rows) else [""] * len(header_names)
                                row_matches: list[bool] = []
                                for j, csv_val in enumerate(csv_row):
                                    ui_val = ui_row[j] if j < len(ui_row) else ""
                                    csv_text = (str(csv_val) or "").strip()
                                    csv_text = csv_text.replace(" ","")
                                    ui_text = (str(ui_val) or "").strip()
                                    ui_text = ui_text.strip().replace(" ","")
                                    is_match = ui_text.lower() in csv_text.lower()

                                    row_matches.append(is_match)
                                    if is_match:
                                        log(
                                            f"✅ Row {i+1}, column '{header_names[j]}' matches: '{csv_text}'"
                                        )
                                    else:
                                        log(
                                            f"❌ Row {i+1}, column '{header_names[j]}' mismatch: CSV='{csv_text}' vs UI='{ui_text}'"
                                        )
                                ui_match_matrix.append(row_matches)
                        else:
                            log("Non-sticket export type; skipping UI comparison for this run.")
                    except Exception as cmp_ex:
                        log(f"⚠️ UI comparison for sticket export failed: {cmp_ex}")
                        ui_match_matrix = None
                    # ------------------------------------------------------------------------------

                    # ---- DELETE CSV FILE AFTER PROCESSING ----
                    try:
                        file_path.unlink()  # delete the CSV
                        log(f"🗑️ Deleted processed file: {file_path}")
                    except Exception as delete_ex:
                        log(f"⚠️ Could not delete CSV file: {delete_ex}")

                    # build sample table HTML fragment from header_names and rows_sample
                    def build_sample_table_html(header_names_, data_rows_, match_matrix_=None):
                        import html as _html
                        parts = [
                            "<div style='margin-top:12px;'>",
                            "<h2 style='margin:6px 0 8px 0;font-size:1.05rem;color:#1f5f0f;'>CSV Sample Rows (filtered)</h2>",
                            "<div style='overflow:auto;max-width:100%'>",
                            "<table class='sample' role='table'>",
                            "<thead><tr>",
                        ]
                        for h in header_names_:
                            parts.append(f"<th>{_html.escape(str(h))}</th>")
                        parts.append("</tr></thead><tbody>")
                        if data_rows_:
                            for i, r in enumerate(data_rows_):
                                parts.append("<tr>")
                                for j, c in enumerate(r):
                                    style_attr = ""
                                    if match_matrix_ is not None and i < len(match_matrix_) and j < len(match_matrix_[i]):
                                        m = match_matrix_[i][j]
                                        if m is True:
                                            # green for match
                                            style_attr = " style='background:#e6ffed;'"
                                        elif m is False:
                                            # red for mismatch
                                            style_attr = " style='background:#ffecec;'"
                                    parts.append(
                                        f"<td{style_attr}>{_html.escape(str(c))}</td>"
                                    )
                                parts.append("</tr>")
                        else:
                            parts.append("<tr><td colspan='100%'>No sample rows captured.</td></tr>")
                        parts.append("</tbody></table></div></div>")
                        return "\n".join(parts)

                    table_html = build_sample_table_html(header_names, rows_sample, ui_match_matrix)

                    # save HTML (once) with the filtered table
                    if not html_report_written:
                        save_logs_to_html(selected_customer, selected_export, filename=LOG_HTML_FILE,
                                          sample_table_html=table_html)
                        log("✅ Inserted filtered CSV sample table into HTML report.")
                else:
                    log(f"❌ Unexpected end status '{status_str}' when percent==100")
                    raise Exception(f"Unexpected end status '{status_str}'")

        except Exception as ex:
            log(f"❌ export_dashboard failed: {ex}")
            raise

        finally:
            try:
                if original_window and original_window in self.driver.window_handles:
                    # nothing extra needed here
                    pass
            except Exception as sw_ex:
                log(f"Notice: failed to access original window in finally: {sw_ex}")


# ─── Export flow (called from UI) ─────────────────────────────────
def run_export_flow(selected_customer: str, selected_export: str, selected_env: str, master: Tk) -> None:
    """
    Run the whole Selenium + validation flow for the given customer/export/env,
    using `master` as the Tk root for ProgressWindow and messageboxes.
    """
    global html_report_written

    try:
        total_steps = 10
        progress = ProgressWindow(master, total_steps)

        c1 = ContactExport(CONFIG_FILE_PATH)

        env_upper = (selected_env or "").upper()
        if env_upper == "CERT":
            c1.certlogin_cozeva(selected_customer, progress)
        elif env_upper == "PROD":
            c1.prodlogin_cozeva(selected_customer, progress)
        else:
            raise RuntimeError(f"Unknown environment: {selected_env!r}")

        if selected_export == "Contact Export":
            c1.contact_export(progress)
        elif selected_export == "Sticket Export":
            c1.sticket_export(progress)
        else:
            messagebox.showwarning("Warning", f"Unknown export option: {selected_export}", parent=master)
            log(f"Unknown export option selected: {selected_export}")

        c1.export_dashboard(selected_customer, selected_export, progress)
        c1.logout_cozeva(progress, customer=selected_customer, export_type=selected_export)

        messagebox.showinfo("Success", "Validation completed successfully!", parent=master)

    except Exception as e:
        log(f"❌ {e}")
        try:
            if not html_report_written:
                save_logs_to_html(
                    selected_customer if selected_customer else "Unknown",
                    selected_export if selected_export else "Unknown"
                )
        except Exception as ex:
            log(f"❌ Failed to write log after exception: {ex}")
        messagebox.showerror("Error", str(e), parent=master)


# Optional: standalone main for testing this file directly
def main() -> None:
    root = Tk()
    root.withdraw()

    from Export_DashboardUI import start_ui  # lazy import to avoid circular issues

    selected_customer, selected_export, selected_env = start_ui(root)
    if not selected_customer or not selected_export or not selected_env:
        log("No selection made. Exiting.")
        root.destroy()
        return

    run_export_flow(selected_customer, selected_export, selected_env, root)

    try:
        root.destroy()
    except Exception:
        pass


if __name__ == "__main__":
    main()
