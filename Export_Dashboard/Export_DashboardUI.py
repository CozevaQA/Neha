import csv
from tkinter import *
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from typing import Union

# ─── Configuration ────────────────────────────────────────────
EXPORT_OPTIONS = ["Select", "Contact Export", "Sticket Export"]
CSV_FILE_PATH = r"C:\Users\nsikder\PycharmProjects\Export Dashboard\Customer.csv"

bg_color = "#7dab41"

# GLOBAL FONT SETTINGS — change here to affect all buttons
BUTTON_FONT = ("Arial", 12, "bold")   # main big buttons
LABEL_FONT = ("Arial", 11, "bold")    # labels for customer/export
SMALL_BUTTON_FONT = ("Arial", 9, "bold")  # for CERT/PROD buttons


# ─── Load Customer CSV ─────────────────────────────────────────
def load_customers_from_csv(filename: str):
    customers = []
    try:
        with open(filename, newline="", encoding="utf-8") as csvfile:
            reader = csv.DictReader(csvfile)
            if "Customer Name" not in reader.fieldnames:
                messagebox.showerror("Error", "CSV must have a 'Customer Name' column.")
                return ["Select"]
            for row in reader:
                if row["Customer Name"]:
                    customers.append(row["Customer Name"])
    except FileNotFoundError:
        messagebox.showerror("Error", f"File '{filename}' not found.")
        return ["Select"]
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return ["Select"]
    return ["Select"] + customers


# ─── Branding (favicon + Cozeva logo bottom-right) ────────────
def apply_branding(window: Union[Tk, Toplevel]) -> None:
    # Favicon
    try:
        window.iconbitmap("favicon.ico")
    except Exception:
        pass

    # Cozeva logo bottom-right
    try:
        logo_image = Image.open("cozeva.png")
        logo_image = logo_image.resize((40, 30), Image.Resampling.LANCZOS)
        logo_photo = ImageTk.PhotoImage(logo_image)
        window._logo_photo = logo_photo  # prevent GC

        logo_label = Label(window, image=logo_photo, borderwidth=0, bg=bg_color)
        logo_label.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)
    except Exception:
        pass


# ─── Second Window UI ───────────────────────────────────────────
def start_ui(parent: Tk):
    """
    Opens the Contact/Sticket Export selection window as a Toplevel
    and returns (selected_customer, selected_export, selected_env).

    selected_env will be "CERT" or "PROD".
    """
    win = Toplevel(parent)
    win.title("Contact/Sticket Export")
    win.geometry("500x320")
    win.configure(bg=bg_color)
    win.resizable(False, False)

    apply_branding(win)

    form_frame = Frame(win, bg=bg_color)
    form_frame.pack(padx=20, pady=20, anchor="w")

    customer_list = load_customers_from_csv(CSV_FILE_PATH)

    # ─── Customer Dropdown ────────────────────────────────
    Label(form_frame, text="Select Customer:", bg=bg_color, fg="white", font=LABEL_FONT).grid(
        row=0, column=0, sticky="w", padx=5, pady=10
    )
    Label(form_frame, text=" *", bg=bg_color, fg="red", font=LABEL_FONT).grid(
        row=0, column=0, sticky="w", padx=(130, 0), pady=10
    )

    customer_var = StringVar()
    customer_dropdown = ttk.Combobox(form_frame, textvariable=customer_var, state="readonly", width=38)
    customer_dropdown["values"] = customer_list
    customer_dropdown.current(0)
    customer_dropdown.grid(row=0, column=1, padx=5, ipady=4, pady=10)

    # ─── Export Options ────────────────────────────────
    Label(form_frame, text="Select Export Type:", bg=bg_color, fg="white", font=LABEL_FONT).grid(
        row=1, column=0, sticky="w", padx=5, pady=10
    )
    Label(form_frame, text=" *", bg=bg_color, fg="red", font=LABEL_FONT).grid(
        row=1, column=0, sticky="w", padx=(145, 0), pady=10
    )

    export_var = StringVar()
    export_dropdown = ttk.Combobox(form_frame, textvariable=export_var, state="readonly", width=25)
    export_dropdown["values"] = EXPORT_OPTIONS
    export_dropdown.current(0)
    export_dropdown.grid(row=1, column=1, padx=5, ipady=4, pady=10)

    # ─── Environment Selection (CERT/PROD) ────────────────────────
    Label(
        form_frame,
        text="Select Environment:",
        bg=bg_color,
        fg="white",
        font=("Arial", 10, "bold")
    ).grid(
        row=2, column=0, sticky="w", padx=5, pady=10
    )

    Label(
        form_frame,
        text=" *",
        bg=bg_color,
        fg="red",
        font=("Arial", 10, "bold")
    ).grid(
        row=2, column=0, sticky="w", padx=(135, 0), pady=10
    )

    selected_env = {"value": None}  # "CERT" / "PROD"

    env_button_frame = Frame(form_frame, bg=bg_color)
    env_button_frame.grid(row=2, column=1, padx=5, pady=10, sticky="w")

    def set_env(env: str):
        selected_env["value"] = env
        # reset styles
        btn_cert.config(bg="white", fg=bg_color, relief="raised")
        btn_prod.config(bg="white", fg=bg_color, relief="raised")
        # highlight selected
        if env == "CERT":
            btn_cert.config(bg=bg_color, fg="white", relief="sunken")
        else:
            btn_prod.config(bg=bg_color, fg="white", relief="sunken")

    btn_cert = Button(
        env_button_frame,
        text="CERT",
        command=lambda: set_env("CERT"),
        bg="white",
        fg=bg_color,
        width=10,
        font=SMALL_BUTTON_FONT,
    )
    btn_cert.pack(side="left", padx=5)

    btn_prod = Button(
        env_button_frame,
        text="PROD",
        command=lambda: set_env("PROD"),
        bg="white",
        fg=bg_color,
        width=10,
        font=SMALL_BUTTON_FONT,
    )
    btn_prod.pack(side="left", padx=5)

    # ─── Selected values to return ────────────────────────────────
    selected_customer = {"value": None}
    selected_export = {"value": None}

    # ─── Bottom "* Mandatory Fields" text ─────────────────────────
    frame_bottom = Frame(win, bg=bg_color)
    frame_bottom.pack(side="bottom", anchor="w", pady=5, padx=10)

    Label(frame_bottom, text="*", fg="red", bg=bg_color, font=("Arial", 9, "italic")).pack(side="left")
    Label(frame_bottom, text="Mandatory Fields", fg="white", bg=bg_color,
          font=("Arial", 8, "italic")).pack(side="left")

    # ─── Submit Button (text) ──────────────────────
    def on_submit():
        cust = customer_var.get()
        exp = export_var.get()

        if cust == "Select":
            messagebox.showwarning("Warning", "Please select a customer.", parent=win)
        elif exp == "Select":
            messagebox.showwarning("Warning", "Please select an export option.", parent=win)
        elif not selected_env["value"]:
            messagebox.showwarning("Warning", "Please select CERT or PROD.", parent=win)
        else:
            selected_customer["value"] = cust
            selected_export["value"] = exp
            win.destroy()

    # Plain Submit button (no GIF)
    submit_button = Button(
        win,
        text="Submit",
        command=on_submit,
        bg="#b2df78",
        fg="black",
        width=20,
        font=BUTTON_FONT,
        padx=15,
        pady=6
    )
    submit_button.pack(pady=12)

    parent.wait_window(win)
    return selected_customer["value"], selected_export["value"], selected_env["value"]


# ─── First Window ────────────────────────────────────────────────
def launch_main_window():
    root = Tk()
    root.title("Choose your option")
    root.geometry("400x230")
    root.configure(bg=bg_color)
    root.resizable(False, False)

    apply_branding(root)

    def contact_log():
        messagebox.showinfo("Contact log", "Contact log clicked!", parent=root)

    def sticket_log():
        messagebox.showinfo("Sticket log", "Sticket log clicked!", parent=root)

    def export_dashboard():
        # Import here to avoid circular imports on module load
        from Export_Functionality import run_export_flow  # ensure filename matches

        # Hide main window while we do the selection + Selenium work
        root.withdraw()

        selected_customer, selected_export, selected_env = start_ui(root)

        # If user closes the window / cancels (without proper selection)
        if not selected_customer or not selected_export or not selected_env:
            root.deiconify()
            return

        # Run Selenium functionality with chosen values
        run_export_flow(selected_customer, selected_export, selected_env, root)

        # When done, close the whole app
        try:
            root.destroy()
        except Exception:
            pass

    btn1 = Button(
        root,
        text="Contact log",
        command=contact_log,
        bg="#b2df78",
        fg="black",
        width=20,
        font=BUTTON_FONT
    )
    btn1.pack(pady=8)

    btn2 = Button(
        root,
        text="Sticket log",
        command=sticket_log,
        bg="#b2df78",
        fg="black",
        width=20,
        font=BUTTON_FONT
    )
    btn2.pack(pady=8)

    btn3 = Button(
        root,
        text="Export Dashboard",
        command=export_dashboard,
        bg="#b2df78",
        fg="black",
        width=20,
        font=BUTTON_FONT
    )
    btn3.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    launch_main_window()
