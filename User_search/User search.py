import csv
import threading
from tkinter import *
from tkinter import ttk, messagebox
from PIL import Image, ImageTk

from user_validation_runner import run_user_validation

# ─── Constants ─────────────────────────────────────────
CSV_FILE_PATH = r"C:\Users\nsikder\PycharmProjects\User search\Customer.csv"
bg_color = "#7dab41"

FONT_BOLD_10 = ("Arial", 10)
FONT_BOLD_11 = ("Arial", 11, "bold")
FONT_BOLD_13 = ("Arial", 13, "bold")
FONT_COMBOBOX = ("Arial", 9, "bold")


# ─── Load Customer CSV ─────────────────────────────────
def load_customers_from_csv(filename: str):
    try:
        with open(filename, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            customers = [r["Customer Name"] for r in reader if r.get("Customer Name")]
        return ["Select"] + customers
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return ["Select"]


# ─── Tooltip ──────────────────────────────────────────
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width() + 10
        y = self.widget.winfo_rooty()
        self.tip = Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        Label(
            self.tip,
            text=self.text,
            bg=bg_color,
            fg="black",
            font=FONT_BOLD_10,
            padx=6,
            pady=4,
            relief="solid",
            borderwidth=1,
        ).pack()

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ─── FAQ Window ───────────────────────────────────────
def open_faq_window(parent):
    faq = Toplevel(parent)
    faq.title("FAQ")
    faq.geometry("350x250")
    faq.configure(bg="white")
    faq.resizable(False, False)

    Label(faq, text="Frequently Asked Questions", font=FONT_BOLD_13).pack(pady=10)

    Label(
        faq,
        text=(
            "• CERT is selected by default\n\n"
            "• Select a customer\n\n"
            "• Choose at least one area\n\n"
            "• Use Select All to toggle options\n\n"
            "• Click SUBMIT to start validation"
        ),
        font=FONT_BOLD_10,
        wraplength=320,
        justify="left",
        bg="white",
    ).pack(padx=15, pady=10)


# ─── Main Window ──────────────────────────────────────
def launch_main_window():
    win = Tk()
    win.title("User Search Validation")
    win.geometry("520x560")
    win.configure(bg="white")
    win.resizable(False, False)

    # ─── FAQ Button ─────────────────────
    faq_btn = Button(
        win,
        text="?",
        font=("Arial", 12, "bold"),
        width=2,
        bg=bg_color,
        fg="white",
        command=lambda: open_faq_window(win),
    )
    faq_btn.place(relx=1, x=-10, y=10, anchor="ne")
    ToolTip(faq_btn, "Help / FAQ")

    # ─── Logo ─────────────────────
    try:
        img = Image.open("cozeva2.png").resize((170, 60), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        win._logo = photo
        Label(win, image=photo, bg="white").pack(pady=(10, 5))
    except Exception:
        pass

    # ─── Customer Dropdown ─────────────────────
    customers = load_customers_from_csv(CSV_FILE_PATH)

    Label(win, text="Select Customer", bg="white", font=FONT_BOLD_11).pack(pady=(5, 3))

    style = ttk.Style()
    style.configure("Big.TCombobox", font=FONT_COMBOBOX, padding=6)
    win.option_add("*TCombobox*Listbox.font", FONT_COMBOBOX)

    customer_var = StringVar()
    customer_cb = ttk.Combobox(
        win,
        textvariable=customer_var,
        values=customers,
        state="readonly",
        width=38,
        style="Big.TCombobox",
    )
    customer_cb.current(0)
    customer_cb.pack(pady=(0, 8))

    # ─── Scrollable Area ─────────────────────
    container = Frame(win, bg="white")
    container.pack(fill="both", expand=True, padx=10)

    canvas = Canvas(container, bg="white", highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    frame = Frame(canvas, bg="white")

    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-event.delta / 120), "units")

    canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
    canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

    areas = [
        "User List", "Batch Share", "Secure Messaging",
        "Analytics (Batch share, workbook share)", "Support Ticket",
        "Case Management", "User Creation", "Provider Creation",
        "Connect Account", "Delete Testing Data", "Log In",
        "Reset Password", "PCR Support Tool", "Bridge", "Add Delegate"
    ]

    tooltip_texts = {
        "User List": "Search and validate users",
        "Batch Share": "Verify batch sharing functionality",
        "Secure Messaging": "Validate secure message access",
        "Analytics (Batch share, workbook share)": "Analytics sharing checks",
        "Support Ticket": "Verify support ticket search",
        "Case Management": "Case management validation",
        "User Creation": "Check newly created users",
        "Provider Creation": "Validate provider search",
        "Connect Account": "Connected account checks",
        "Delete Testing Data": "Ensure deleted data is not searchable",
        "Log In": "Login-related search validation",
        "Reset Password": "Password reset audit validation",
        "PCR Support Tool": "PCR support tool access",
        "Bridge": "Bridge module validation",
        "Add Delegate": "Delegated user validation",
    }

    vars_map = {}

    # ✅ SELECT ALL
    select_all_var = BooleanVar(value=True)

    def toggle_all():
        for v in vars_map.values():
            v.set(select_all_var.get())

    select_all_cb = Checkbutton(
        frame,
        text="Select All",
        variable=select_all_var,
        command=toggle_all,
        bg="white",
        font=("Arial", 10, "bold"),
    )
    select_all_cb.pack(anchor="w", pady=(2, 6))
    ToolTip(select_all_cb, "Select / deselect all areas")

    # Individual areas
    for area in areas:
        var = BooleanVar(value=True)
        cbx = Checkbutton(frame, text=area, variable=var,
                          bg="white", font=FONT_BOLD_10)
        cbx.pack(anchor="w", pady=2)
        ToolTip(cbx, tooltip_texts.get(area, area))
        vars_map[area] = var

    # ─── Environment Selection ─────────────────────
    env_var = StringVar(value="CERT")

    env_frame = Frame(win, bg="white")
    env_frame.pack(pady=(0, 10))

    Label(env_frame, text="Environment:", bg="white", font=FONT_BOLD_11) \
        .pack(side="left", padx=(0, 10))

    Radiobutton(env_frame, text="CERT", variable=env_var, value="CERT",
                bg="white", font=FONT_BOLD_10).pack(side="left", padx=5)

    Radiobutton(env_frame, text="PROD", variable=env_var, value="PROD",
                bg="white", font=FONT_BOLD_10).pack(side="left", padx=5)

    # ─── Submit ─────────────────────
    def submit():
        if customer_var.get() == "Select":
            messagebox.showwarning("Validation", "Please select a customer.")
            return

        selected = [k for k, v in vars_map.items() if v.get()]
        if not selected:
            messagebox.showwarning("Validation", "Please select at least one area.")
            return

        win.withdraw()

        threading.Thread(
            target=run_user_validation,
            args=(win, customer_var.get(), selected, env_var.get()),
            daemon=True,
        ).start()

    Button(
        win,
        text="SUBMIT",
        bg=bg_color,
        fg="white",
        font=FONT_BOLD_11,
        width=22,
        command=submit,
    ).pack(pady=10)

    win.mainloop()


if __name__ == "__main__":
    launch_main_window()
