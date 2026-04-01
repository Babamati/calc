
import tkinter as tk
from tkinter import ttk, filedialog
from datetime import datetime
from fpdf import FPDF
import math
import calendar
from dateutil.relativedelta import relativedelta
from tkcalendar import DateEntry
from tkinter import StringVar
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np
from decimal import Decimal
import os
import platform
import tempfile
from tkinter import messagebox

# Tooltip class
class CreateToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, _cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "12", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
        self.tipwindow = None

        
# Financial formulas
def compute_macaulay_duration(cashflows, discount_factors, times, price):
    return sum(t * cf * df for t, cf, df in zip(times, cashflows, discount_factors)) / price

def compute_convexity(cashflows, times, ytm, price, freq, is_bill):
    if is_bill:
        # For a bill, t is a single year fraction, so just apply the standard formula
        t = times[0]
        return (t * (t + 1)) / ((1 + ytm) ** 2)
    else:
        # For coupon bonds, convert t to periods and apply full formula with annualization
        y_period = ytm / freq
        convexity_sum = sum(
            cf * t * (t + 1) / (1 + y_period) ** (t + 2)
            for cf, t in zip(cashflows, [t * freq for t in times])
        )
        return convexity_sum / (price * freq ** 2)

def get_last_coupon_date(settlement, freq, maturity):
    delta_months = 12 // freq
    anchor_day = maturity.day
    anchor_month = maturity.month

    # start stepping backward from maturity until we find the last coupon before or on settlement
    last_coupon = maturity
    while last_coupon > settlement:
        last_coupon = add_months(last_coupon, -delta_months)
        # force same day-of-month as anchor
        try:
            last_coupon = last_coupon.replace(day=anchor_day)
        except ValueError:
            # handle month-end rollover (e.g., 30 Feb doesn't exist)
            last_coupon = last_coupon.replace(day=1) + relativedelta(months=1) - relativedelta(days=1)

    return last_coupon
            
def get_coupon_schedule(maturity, freq, settlement):
    delta_months = 12 // freq
    coupon_dates = []
    next_coupon = maturity
    while next_coupon > settlement:
        coupon_dates.insert(0, next_coupon)
        next_coupon = add_months(next_coupon, -delta_months)
    return coupon_dates

def calculate_accrued_interest(settlement, maturity, coupon_amount, freq, ex_days, convention):
    last_coupon = get_last_coupon_date(settlement, freq, maturity)
    next_coupon = add_months(last_coupon, 12 // freq)

    # Check if we're in the ex-interest period
    in_ex_interest = (next_coupon - settlement).days <= ex_days

    if in_ex_interest:
        # Shift the coupon schedule forward
        last_coupon = next_coupon
        next_coupon = add_months(last_coupon, 12 // freq)

    # Compute accrued and period days based on convention
    if convention == "30/360":
        d1 = min(settlement.day, 30)
        d2 = min(last_coupon.day, 30)
        accrued_days = (settlement.year - last_coupon.year) * 360 + \
                       (settlement.month - last_coupon.month) * 30 + (d1 - d2)
        period_days = 360 // freq
    
    elif convention == "ACT/360":
        accrued_days = (settlement - last_coupon).days
        period_days = 360 // freq
      
    else:
        accrued_days = (settlement - last_coupon).days
        period_days = (next_coupon - last_coupon).days

    if period_days == 0:
        return Decimal("0.0")

    return Decimal(coupon_amount) * Decimal(accrued_days) / Decimal(period_days)

    
# Define navigation logic first
def display_amort_page(page):
    global current_page
    amort_text.delete("1.0", tk.END)

    # Header always shown
    amort_text.insert(tk.END, f"{'Date':15} {'Cashflow':15} {'Discount Factor':18} {'Present Value':18}\n")
    amort_text.insert(tk.END, "=" * 70 + "\n")

    start = page * page_size
    end = start + page_size
    for line in amort_lines[start:end]:
        amort_text.insert(tk.END, line + "\n")

    current_page = page
    update_nav_buttons()

def update_nav_buttons():
    total_pages = (len(amort_lines) + page_size - 1) // page_size
    prev_button.config(state="normal" if current_page > 0 else "disabled")
    next_button.config(state="normal" if current_page < total_pages - 1 else "disabled")
    page_label.config(text=f"Page {current_page + 1} of {total_pages}")

def normalize_rate_field(ent):
    try:
        val = float(ent.get().replace("%", "").strip())
        ent.delete(0, tk.END)
        ent.insert(0, f"{val:.6f}%")
    except:
        ent.delete(0, tk.END)

def normalize_currency_field(ent):
    try:
        val_str = ent.get().replace("$", "").replace(",", "").strip().lower()
        val_str = val_str.replace(" ", "")

        if val_str.endswith("k"):
            val = float(val_str[:-1]) * 1_000
        elif val_str.endswith("m"):
            val = float(val_str[:-1]) * 1_000_000
        elif val_str.endswith("b"):
            val = float(val_str[:-1]) * 1_000_000_000
        else:
            val = float(val_str)

        ent.delete(0, tk.END)
        ent.insert(0, f"${val:,.2f}")
    except:
        ent.delete(0, tk.END)
# Flag to prevent multiple messageboxes at the same time
popup_active = False

def validate_numeric_popup(entry_widget, label_text):
    global popup_active
    try:
        text = entry_widget.get().replace("%", "").strip()
        float(text)
        return True
    except ValueError:
        if not popup_active:
            popup_active = True
            messagebox.showerror("Invalid Input", f"{label_text} must be a valid number.")
            popup_active = False
        
        return False

AUTO_DEBOUNCE_DELAY = 300  # milliseconds
auto_calc_enabled = False
calculate_scheduled = None

def try_auto_calculate_debounced():
    global calculate_scheduled
    if not auto_calc_enabled:
        return  # prevent auto-calc if disabled
    if calculate_scheduled:
        root.after_cancel(calculate_scheduled)
    calculate_scheduled = root.after(AUTO_DEBOUNCE_DELAY, calculate)

# Helper function to show Calendar popup for an Entry field
from tkcalendar import Calendar

def show_calendar_for_entry(entry):
    def on_date_selected(event):
        date = cal.selection_get()
        entry.delete(0, tk.END)
        entry.insert(0, date.strftime("%d/%m/%Y"))
        top.destroy()
        # Trigger auto-calculate if enabled
        if auto_calc_enabled:
            try_auto_calculate_debounced()

    # Create popup window
    top = tk.Toplevel(root)
    top.withdraw()
    top.transient(root)
    top.grab_set()
    top.title("Select Date")

    # Try to parse existing date in Entry field
    try:
        current_text = entry.get().strip()
        init_date = datetime.strptime(current_text, "%d/%m/%Y")
        year = init_date.year
        month = init_date.month
        day = init_date.day
    except Exception:
        # If parsing fails, fallback to today
        today = datetime.today()
        year = today.year
        month = today.month
        day = today.day

    # Create calendar with initial date
    cal = Calendar(top, date_pattern='dd/MM/yyyy',
                   year=year, month=month, day=day)
    cal.pack(padx=10, pady=10)

    # Reliable binding with small delay
    cal.bind("<<CalendarSelected>>", lambda event: top.after(10, lambda: on_date_selected(event)))

    # Improve sizing — update geometry and center the popup
    cal.update_idletasks()

    popup_width = cal.winfo_reqwidth() + 20
    popup_height = cal.winfo_reqheight() + 20

    min_width = 300
    min_height = 300
    popup_width = max(popup_width, min_width)
    popup_height = max(popup_height, min_height)

    x = root.winfo_x() + (root.winfo_width() // 2) - (popup_width // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (popup_height // 2)
    top.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
    top.deiconify() 
# GUI setup
root = tk.Tk()
root.title("Bond Calculator")
root.resizable(False, False)  # Prevent resizing

amort_lines = []
page_size = 8
current_page = 0

# === Window sizing ===
# === Window sizing ===
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

base_width = 915
base_height = 870

margin_x = 50
margin_y = 30

max_width = screen_width - 2 * margin_x
max_height = screen_height - 2 * margin_y

target_width = base_width if base_width <= max_width else max_width
target_height = base_height if base_height <= max_height else max_height

x_position = (screen_width - target_width) // 2
y_position = 30  # reduce top margin — looks better

if screen_width >= base_width and screen_height >= base_height:
    root.geometry(f"{base_width}x{base_height}+{x_position}+{y_position}")
else:
    # Use smaller window with scrolling
    root.geometry(f"{max_width}x{max_height}+{x_position}+{y_position}")
#root.minsize(base_width, base_height)
root.resizable(True, True)

# === ScrollableFrame class ===
class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)

        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

# === Wrapper frame ===
wrapper_frame = ScrollableFrame(root)
wrapper_frame.pack(fill="both", expand=True)

# === Main content frame inside ScrollableFrame ===
main_frame = tk.Frame(wrapper_frame.scrollable_frame, width=base_width, height=base_height)
main_frame.pack_propagate(False)
main_frame.pack()

# === Heading section ===
heading_frame = tk.Frame(main_frame)
heading_frame.pack(pady=(10, 0),fill="x" )

tk.Label(
    heading_frame,
    text="NORTHERN TERRITORY TREASURY CORPORATION",
    font=("Arial", 14, "bold"),
    anchor="center"
).pack()

tk.Label(
    heading_frame,
    text="BOND PRICE CALCULATOR",
    font=("Arial", 14, "bold"),
    anchor="center"
).pack(pady=(5, 15))

# === Input tabs ===
input_tabs = ttk.Notebook(main_frame)
input_tabs.pack(fill="both", expand=True, padx=10, pady=10)

# === Bond Inputs tab ===
bond_input_tab = ttk.LabelFrame(main_frame, text="Bond Inputs")
bond_input_tab.configure(padding=10, relief="groove", borderwidth=2)
bond_input_tab.pack(fill="both", expand=True, padx=10, pady=10)

# Frame to hold the chart just below bond inputs
chart_frame = tk.Frame(bond_input_tab)
chart_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

# LEFT and RIGHT column layout for bond input tab
left_frame = tk.Frame(bond_input_tab)
left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=5)

right_frame = tk.Frame(bond_input_tab)
right_frame.pack(side="left", fill="both", expand=True, padx=10, pady=5)

# === LEFT column inputs ===
left_inputs = [
    ("Settlement Date", DateEntry, "date"),
    ("Maturity Date", DateEntry, "date"),
    ("Yield (% per annum)", ttk.Entry, "rate3"),
    ("Coupon Rate (% of face value)", ttk.Entry, "rate3"),
    ("Face Value ($)", ttk.Entry, "currency")
]
entry_vars = []

for i, (label_text, widget_type, fmt) in enumerate(left_inputs):
    tk.Label(left_frame, text=label_text).grid(row=i, column=0, sticky="w")

    if fmt == "date":
        date_frame = tk.Frame(left_frame)
        entry = ttk.Entry(date_frame, width=23)
        calendar_button = ttk.Button(date_frame, text="📅", width=3,
                                     command=lambda e=entry: show_calendar_for_entry(e))

        entry.pack(side="left", padx=(0, 2))
        calendar_button.pack(side="left")

        date_frame.grid(row=i, column=1, padx=5, pady=2, sticky="w")

    else:
        entry = widget_type(left_frame, width=30)

        if fmt == "rate3":
            entry.bind("<FocusOut>", lambda e, ent=entry: normalize_rate_field(ent))
        elif fmt == "currency":
            entry.bind("<FocusOut>", lambda e, ent=entry: normalize_currency_field(ent))

        entry.grid(row=i, column=1, padx=5, pady=2)

    # Tooltip setup
    tooltip_text = {
        "Settlement Date": "Date when the bond is settled or purchased. dd/MM/yyyy. ",
        "Maturity Date": "Date when the bond matures and principal is repaid. dd/MM/yyyy.",
        "Yield (% per annum)": "Annual yield expected by the investor. E.g. 4.250. Will auto-format",
        "Coupon Rate (% of face value)": "Annual coupon rate as % of face value. E.g. 5.000. Will auto format",
        "Face Value ($)": "Enter amount, e.g. 100000 or 50k, 1.5m, 2b. Will auto-format."
    }.get(label_text, "Enter a value.")

    CreateToolTip(entry, tooltip_text)

    # Numeric popup validation for rates
    if label_text in ["Yield (% per annum)", "Coupon Rate (% of face value)"]:
        entry.bind("<FocusOut>", lambda e, ent=entry, label=label_text: validate_numeric_popup(ent, label))

    entry_vars.append(entry)

# === RIGHT column inputs ===
coupon_frequency_var = tk.StringVar(value="1")
interest_status_var = tk.StringVar(value="Cum")
day_count_var = tk.StringVar(value="ACT/ACT")
ex_interest_days_var = tk.IntVar(value=7)

right_inputs = [
    ("Coupon Frequency", ttk.Combobox, coupon_frequency_var, ["1", "2", "4"]),
    ("Interest Status", ttk.Combobox, interest_status_var, ["Cum", "Ex"]),
    ("Day Count Convention", ttk.Combobox, day_count_var, ["ACT/ACT", "30/360", "ACT/360", "ACT/365"]),
    ("Ex-Interest Period (days)", ttk.Spinbox, ex_interest_days_var, list(range(0, 31)))
]

for i, (label_text, widget_type, var, values) in enumerate(right_inputs):
    tk.Label(right_frame, text=label_text).grid(row=i, column=0, sticky="w")

    if widget_type == ttk.Spinbox:
        widget = widget_type(right_frame, from_=0, to=30, textvariable=var, width=28)
    else:
        widget = widget_type(right_frame, textvariable=var, values=values, width=28)

    widget.grid(row=i, column=1, padx=5, pady=2)

    if label_text == "Day Count Convention":
        CreateToolTip(widget, "Select the day count basis for interest accrual.")
    elif label_text == "Ex-Interest Period (days)":
        CreateToolTip(widget, "How many days before coupon date bond enters ex-interest.\nDefault: 7.")
    elif label_text == "Coupon Frequency":
        CreateToolTip(widget, "How often coupons are paid.\n1 = Annual, 2 = Semi-Annual, 4 = Quarterly")

# === Info message below right column ===
info_message_label = tk.Label(
    right_frame,
    text="",
    font=("Arial", 10, "bold"),
    fg="red",
    anchor="w",
    justify="left"
)
info_message_label.grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=5)

# Bind inputs after UI is initialized


# Rounding Option Frame
rounding_var = tk.StringVar(value="3dp")
#rounding_var.trace_add("write", lambda *args: try_auto_calculate_debounced())

rounding_frame = tk.LabelFrame(right_frame, text="Rounding", padx=5, pady=5)
rounding_frame.grid(row=0, column=2, rowspan=5, sticky="nw", padx=(10, 0))

tk.Radiobutton(rounding_frame, text="12dp", variable=rounding_var, value="12dp").pack(anchor="w")
tk.Radiobutton(rounding_frame, text="3dp", variable=rounding_var, value="3dp").pack(anchor="w")
tk.Radiobutton(rounding_frame, text="Quantum", variable=rounding_var, value="quantum").pack(anchor="w")

def attach_auto_calculate():
    # Bind left frame entries
    for entry in entry_vars:
        entry.bind("<FocusOut>", lambda e: root.after(10, try_auto_calculate_debounced))

    # Bind right side variable traces
    coupon_frequency_var.trace_add("write", lambda *args: try_auto_calculate_debounced())
    interest_status_var.trace_add("write", lambda *args: try_auto_calculate_debounced())
    day_count_var.trace_add("write", lambda *args: try_auto_calculate_debounced())
    ex_interest_days_var.trace_add("write", lambda *args: try_auto_calculate_debounced())
    rounding_var.trace_add("write", lambda *args: try_auto_calculate_debounced())

def clear_all():
    global auto_calc_enabled
    auto_calc_enabled = False  # prevent auto-calculate immediately after clearing

    for entry in entry_vars:
        entry.delete(0, tk.END)
    for label in price_labels + settlement_labels + risk_labels:
        label.config(text="")
    amort_text.delete("1.0", tk.END)
    info_message_label.config(text="")  # Clears the red info message

# === Actions Section ===
actions_frame = tk.LabelFrame(main_frame, text="Actions")
actions_frame.pack(fill="x", padx=10, pady=5)

# Inner frame to center the button group
button_group = tk.Frame(actions_frame)
button_group.pack(fill="x", expand=True, pady=5)

# Buttons in desired order with equal spacing
calculate_button = tk.Button(button_group, text="Calculate")
calculate_button.pack(side="left", padx=15, expand=True, fill="x")

generate_button = tk.Button(button_group, text="Generate Chart")
generate_button.pack(side="left", padx=15, expand=True, fill="x")

export_button = tk.Button(button_group, text="Export to PDF")
export_button.pack(side="left", padx=15, expand=True, fill="x")

clear_button = tk.Button(button_group, text="Clear")
clear_button.pack(side="left", padx=15, expand=True, fill="x")
clear_button.config(command=clear_all)

output_frame = tk.LabelFrame(main_frame, text="Results")
output_frame.pack(fill="both", expand=True, padx=10, pady=10)

price_frame = tk.LabelFrame(output_frame, text="Price per $100")
price_frame.grid(row=0, column=0, padx=20, pady=5, sticky="n")

settlement_frame = tk.LabelFrame(output_frame, text="Settlement")
settlement_frame.grid(row=0, column=1, padx=20, pady=5, sticky="n")

risk_frame = tk.LabelFrame(output_frame, text="Risk Metrics")
risk_frame.grid(row=0, column=2, padx=20, pady=5, sticky="n")

bold_font = ("Arial", 12, "bold")
monospace_font = ("Arial", 12, "bold")

price_labels = []
settlement_labels = []
risk_labels = []

for i, label in enumerate(["Capital Price", "Accrued Interest", "Gross Price"]):
    tk.Label(price_frame, text=label + ":", anchor="w", justify="left").grid(row=i, column=0, sticky="w")
    val = tk.Label(price_frame, text="", font=monospace_font, anchor="e", justify="right", width=10)
    val.grid(row=i, column=1, sticky="e")
    price_labels.append(val)

for i, label in enumerate(["Capital Amount", "Accrued Interest Amount", "Total Settlement Value"]):
    tk.Label(settlement_frame, text=label + ":", anchor="w", justify="left").grid(row=i, column=0, sticky="w")
    val = tk.Label(settlement_frame, text="", font=monospace_font, anchor="e", justify="right", width=15)
    val.grid(row=i, column=1, sticky="e")
    settlement_labels.append(val)

for i, label in enumerate(["Duration", "Modified Duration", "Convexity"]):
    tk.Label(risk_frame, text=label + ":", anchor="w", justify="left").grid(row=i, column=0, sticky="w")
    val = tk.Label(risk_frame, text="", font=monospace_font, anchor="e", justify="right", width=10)
    val.grid(row=i, column=1, sticky="e")
    risk_labels.append(val)

separator = tk.Frame(main_frame, height=2, bd=1, relief="sunken")
separator.pack(fill="x", padx=5, pady=10)

amort_frame = tk.LabelFrame(main_frame, text="Amortization Schedule")
amort_frame.pack(fill="both", expand=True, padx=10, pady=10)

amort_text = tk.Text(amort_frame, height=10, wrap="none")
amort_text.pack(fill="both", expand=True)

# Add navigation buttons below amort_text
nav_frame = tk.Frame(amort_frame)
nav_frame.pack(pady=5)

prev_button = tk.Button(nav_frame, text="← Previous", command=lambda: display_amort_page(current_page - 1))
prev_button.pack(side="left", padx=10)

next_button = tk.Button(nav_frame, text="Next →", command=lambda: display_amort_page(current_page + 1))
next_button.pack(side="left", padx=10)

page_label = tk.Label(nav_frame, text="Page 1 of 1", font=("Arial", 10, "bold"))
page_label.pack(side="left", padx=10)


def add_months(source_date, months):
    return source_date + relativedelta(months=months)

def year_fraction(start_date, end_date, convention):
    days = (end_date - start_date).days
    if convention == "30/360":
        d1 = start_date.day
        d2 = end_date.day
        m1 = start_date.month
        m2 = end_date.month
        y1 = start_date.year
        y2 = end_date.year
        d1 = min(d1, 30)
        d2 = min(d2, 30) if d1 == 30 else d2
        return ((360 * (y2 - y1)) + (30 * (m2 - m1)) + (d2 - d1)) / 360
    elif convention == "ACT/360":
        return days / 360
    elif convention == "ACT/365":
        return days / 365
    elif convention == "ACT/ACT":
        yf = 0.0
        current = start_date
        while current < end_date:
            year_end = min(datetime(current.year + 1, 1, 1), end_date)
            year_days = 366 if calendar.isleap(current.year) else 365
            yf += (year_end - current).days / year_days
            current = year_end
        return yf
    else:
        raise ValueError("Unsupported day count convention")

def year_fraction_ACT_ACT(start_date, end_date):
    current = start_date
    yf = 0.0
    while current < end_date:
        year_end = min(datetime(current.year + 1, 1, 1), end_date)
        year_days = 366 if calendar.isleap(current.year) else 365
        yf += (year_end - current).days / year_days
        current = year_end
    return yf

def calculate(manual=False):
    # Exit early if required inputs are missing
    values = [ent.get().strip() for ent in entry_vars]
    if not all(values[:5]):  # Check all 5 left-column fields
        if manual:
            messagebox.showerror("Missing Input", "Please fill in all bond input fields before calculating.")
        return
    
    # Manually apply formatting before reading input
    normalize_currency_field(entry_vars[4])  # Face Value
    normalize_rate_field(entry_vars[2])      # Yield
    normalize_rate_field(entry_vars[3])      # Coupon Rate
    
    try:
        settlement = datetime.strptime(entry_vars[0].get(), "%d/%m/%Y")
        maturity = datetime.strptime(entry_vars[1].get(), "%d/%m/%Y")
        # Validation: Settlement must be before Maturity
        if settlement >= maturity:
            # Clear old results
            for lbl in price_labels + settlement_labels + risk_labels:
                lbl.config(text="")
            amort_text.delete("1.0", tk.END)
            page_label.config(text="Page 1 of 1")
            prev_button.config(state="disabled")
            next_button.config(state="disabled")
            
            info_message_label.config(
                text="Error: Settlement Date must be before Maturity Date.",
                fg="red"
            )
            return
        ytm = float(entry_vars[2].get().replace("%", "").strip()) / 100
        face_value = float(entry_vars[4].get().replace("$", "").replace(",", ""))
        coupon_rate = float(entry_vars[3].get().replace("%", "").strip()) / 100
        freq = int(coupon_frequency_var.get())
        convention = day_count_var.get()
        ex_days = int(ex_interest_days_var.get())

        is_ex_interest_selected = interest_status_var.get() == "Ex"

        coupon_amount = (coupon_rate * face_value) / freq
        delta_months = 12 // freq

        # Generate full coupon schedule
        payment_dates = []
        next_coupon = maturity
        while True:
            payment_dates.insert(0, next_coupon)
            next_coupon = add_months(next_coupon, -delta_months)
            if next_coupon <= settlement:
                break

        if payment_dates[0] <= settlement:
            payment_dates = payment_dates[1:]

        price = 0.0
        times, cashflows, discount_factors = [], [], []

        # Determine if it's a bill (single payment)
        is_bill = len(payment_dates) == 1 and payment_dates[0] == maturity
        days_to_next = (payment_dates[0] - settlement).days

        if is_bill:
            if convention == "30/360":
                short_t = days_to_next / 360
            elif convention == "ACT/360":
                short_t = days_to_next / 360
            elif convention == "ACT/365":
                short_t = days_to_next / 365
            else:
                short_t = year_fraction_ACT_ACT(settlement, payment_dates[0])
        else:
            anchor_day = payment_dates[0].day
            last_coupon = add_months(payment_dates[0], -delta_months)
            try:
                last_coupon = last_coupon.replace(day=anchor_day)
            except ValueError:
                last_coupon = last_coupon.replace(day=1) + relativedelta(months=1) - relativedelta(days=1)

            full_coupon_period_days = (payment_dates[0] - last_coupon).days
            if full_coupon_period_days == 0:
                raise ValueError("Coupon period length is zero — check frequency and coupon dates.")

            short_t = days_to_next / full_coupon_period_days

        if is_bill:
            short_df = 1 / (1 + ytm * short_t)
        else:
            short_df = (1 / (1 + ytm / freq)) ** short_t
        
        global amort_lines
        amort_lines = []  # just data lines (headers go in display_amort_page)

        for i, dt in enumerate(payment_dates):
            n = i
            if is_bill:
                cf = coupon_amount + face_value
                df = short_df
                pv = cf * df
                label = "*** FINAL PAYMENT *** "
                times.append(short_t)
            else:
                df_long = 1 / (1 + ytm / freq) ** n
                df = df_long * short_df

                cf = coupon_amount
                if dt == maturity:
                    cf += face_value

                pv = cf * df
                label = "*** FINAL PAYMENT *** " if dt == maturity else ""

                        
            if i == 0 and is_ex_interest_selected and not is_bill:
                label += " (EXCLUDED FROM PRICE)"
            else:
                price += pv

            times.append(n / freq)
            cashflows.append(cf)
            discount_factors.append(df)

            amort_lines.append(f"{dt.strftime('%d/%m/%Y'):15} {cf:15,.2f} {df:18,.10f} {pv:18,.2f} {label}")
            

        # Accrued interest calculation (same for both Cum and Ex)
               # Accrued interest
        if is_ex_interest_selected:
            accrued_interest = calculate_accrued_interest(settlement, maturity, coupon_amount, freq, ex_days, convention) - Decimal(coupon_amount)
        else:
            accrued_interest = calculate_accrued_interest(settlement, maturity, coupon_amount, freq, ex_days, convention) 

        # Full precision price-per-100 values
        gross_price_per_100 = Decimal(price) / Decimal(face_value) * Decimal(100)
        accrued_per_100 = Decimal(accrued_interest) / Decimal(face_value) * Decimal(100)
        clean_price_per_100 = gross_price_per_100 - accrued_per_100
        # Get rounding mode

        rounding_mode = rounding_var.get()

        # Compute full-precision price per 100
        gross_price_per_100 = Decimal(price) / Decimal(face_value) * Decimal(100)
        accrued_per_100 = Decimal(accrued_interest) / Decimal(face_value) * Decimal(100)
        clean_price_per_100 = gross_price_per_100 - accrued_per_100

        # Determine values used for settlement
        if rounding_mode == "quantum":
            if is_bill:
                price_for_settlement = round(gross_price_per_100, 12)
            else:
                price_for_settlement = round(gross_price_per_100, 3)
            accrued_per_100_rounded = round(accrued_per_100, 3)

        elif rounding_mode == "3dp":
            price_for_settlement = round(gross_price_per_100, 3)
            accrued_per_100_rounded = round(accrued_per_100, 3)

        elif rounding_mode == "12dp":
            price_for_settlement = gross_price_per_100  # full precision
            accrued_per_100_rounded = accrued_per_100   # full precision

        else:
            price_for_settlement = round(gross_price_per_100, 3)
            accrued_per_100_rounded = round(accrued_per_100, 3)

        # Compute dollar amounts (always from per-100 values)
        accrued_interest_amount = accrued_per_100_rounded * Decimal(face_value) / Decimal(100)
        total_settlement = price_for_settlement * Decimal(face_value) / Decimal(100)

        # Convert to display-safe strings (rounded)
        display_total_str = f"{total_settlement:,.2f}"
        display_accrued_str = f"{accrued_interest_amount:,.2f}"
        capital_amount = total_settlement - accrued_interest_amount
        display_capital_str = f"{capital_amount:,.2f}"

        # Visual check (sum as floats)
        visual_sum_check = f"{float(display_capital_str.replace(',', '')) + float(display_accrued_str.replace(',', '')):,.2f}" == display_total_str

        # If mismatch, force capital to reconcile visually
        if not visual_sum_check:
            capital_amount = total_settlement - accrued_interest_amount
            display_capital_str = f"{float(display_total_str.replace(',', '')) - float(display_accrued_str.replace(',', '')):,.2f}"


        # Display (always show per-100 values to 3dp, dollar amounts to 2dp)
        # Round per-100 components for display
        clean_str = f"{gross_price_per_100 - accrued_per_100_rounded:.3f}"
        accrued_str = f"{accrued_per_100_rounded:.3f}"
        gross_str = f"{gross_price_per_100:.3f}"

        # Check for visual mismatch
        visual_price_check = f"{float(clean_str) + float(accrued_str):.3f}" == gross_str

        # If mismatch, override clean price to ensure consistency
        if not visual_price_check:
            clean_str = f"{float(gross_str) - float(accrued_str):.3f}"

        # Update per-100 display labels
        price_labels[0].config(text=clean_str)         # Capital Price (display-safe)
        price_labels[1].config(text=accrued_str)       # Accrued Interest
        price_labels[2].config(text=gross_str)         # Gross Price

        settlement_labels[0].config(text=f"${display_capital_str}")
        settlement_labels[1].config(text=f"${display_accrued_str}")
        settlement_labels[2].config(text=f"${display_total_str}")
         
        # Risk metrics
        # Rebuild full cashflows/times/discount_factors for risk metrics (always include all)

        # --- Build full list for risk metrics ---
        risk_cashflows = []
        risk_times = []
        risk_discount_factors = []
        
        for i, dt in enumerate(payment_dates):
            # Time in years to each payment date
            t = year_fraction(settlement, dt, convention) if not is_bill else short_t
        
            # Cashflow amount
            cf = coupon_amount if not is_bill else coupon_amount + face_value
            if dt == maturity:
                cf += face_value if not is_bill else 0
        
            # Discount factor (with fractional compounding)
            df = (1 / (1 + ytm / freq) ** (t * freq)) if not is_bill else short_df
        
            # Append full values — do not exclude any payment
            risk_cashflows.append(cf)
            risk_times.append(t)
            risk_discount_factors.append(df)
        
        # --- Compute dirty price from full risk stream ---
        dirty_price = sum(cf * df for cf, df in zip(risk_cashflows, risk_discount_factors))
        
        # --- Risk Metrics (always based on full dirty price) ---
        macaulay_duration = compute_macaulay_duration(risk_cashflows, risk_discount_factors, risk_times, dirty_price)
        modified_duration = macaulay_duration / (1 + ytm / freq)
        convexity = compute_convexity(risk_cashflows, risk_times, ytm, dirty_price, freq, is_bill)
        
        risk_labels[0].config(text=f"{macaulay_duration:.3f}")
        risk_labels[1].config(text=f"{modified_duration:.3f}")
        risk_labels[2].config(text=f"{convexity:.3f}")

        # Info messages
        in_ex_interest_period = (payment_dates[0] - settlement).days <= ex_days
        
        if is_bill:
            info_message_label.config(text="PRICED AS A BILL", fg="red")
        elif in_ex_interest_period:
            info_message_label.config(
                text="(Ex-Interest on old issues / set new issues to cum interest)",
                fg="red"
            )
        else:
            info_message_label.config(text="")

        entry_vars[0].update()
        entry_vars[1].update()
        display_amort_page(0)
        update_nav_buttons()

    except Exception as e:
        for lbl in price_labels + settlement_labels + risk_labels:
            lbl.config(text="")
        amort_text.delete("1.0", tk.END)
        page_label.config(text="Page 1 of 1")
        prev_button.config(state="disabled")
        next_button.config(state="disabled")
    
        # Use messagebox for long or critical errors
        messagebox.showerror("Calculation Error", str(e))
        info_message_label.config(text="An error occurred. See details in popup.", fg="red")

calculate_button.config(command=calculate)
root.bind('<Return>', lambda event: calculate())

def generate_price_yield_chart():
    from scipy.interpolate import make_interp_spline

    try:
        ytm = float(entry_vars[2].get().replace("%", "").strip()) / 100
        face_value = float(entry_vars[4].get().replace("$", "").replace(",", ""))
        settlement = datetime.strptime(entry_vars[0].get(), "%d/%m/%Y")
        maturity = datetime.strptime(entry_vars[1].get(), "%d/%m/%Y")
        convention = day_count_var.get()
        freq = int(coupon_frequency_var.get())
        ex_days = int(ex_interest_days_var.get())
        is_ex_interest_selected = interest_status_var.get() == "Ex"

        capital_amt = float(settlement_labels[0].cget("text").replace("$", "").replace(",", ""))
        accrued = float(settlement_labels[1].cget("text").replace("$", "").replace(",", ""))
        total_amt = float(settlement_labels[2].cget("text").replace("$", "").replace(",", ""))
        dura = float(risk_labels[0].cget("text"))
        mdura = float(risk_labels[1].cget("text"))
        conv = float(risk_labels[2].cget("text"))

        coupon_rate = float(entry_vars[3].get().replace("%", "").strip()) / 100
        coupon_amt = (coupon_rate * face_value) / freq
        coupon_dates = get_coupon_schedule(maturity, freq, settlement)
        is_bill = len(coupon_dates) == 1 and coupon_dates[0] == maturity

    except Exception as e:
        print("Error parsing values for chart:", e)
        return

    def bond_pv(y):
        if is_bill:
            t = year_fraction(settlement, maturity, convention)
            df = 1 / (1 + y * t)
            pv = (face_value + coupon_amt) * df
            macaulay = t
            modified = macaulay / (1 + y)
            convexity = (t * (t + 1)) / ((1 + y) ** 2)
            return pv, macaulay, modified, convexity
    
        pv = 0.0
        dirty_pv = 0.0
        cashflows, times = [], []
        risk_cf, risk_t, risk_df = [], [], []
    
        for i, dt in enumerate(coupon_dates):
            if dt <= settlement:
                continue
    
            t = year_fraction(settlement, dt, convention)
            cf = coupon_amt
            if dt == maturity:
                cf += face_value
    
            df = 1 / (1 + y / freq) ** (t * freq)
            risk_cf.append(cf)
            risk_t.append(t)
            risk_df.append(df)
            dirty_pv += cf * df
    
            if i == 0 and is_ex_interest_selected:
                continue
    
            pv += cf * df
            times.append(t)
            cashflows.append(cf)
    
        try:
            macaulay = compute_macaulay_duration(risk_cf, risk_df, risk_t, dirty_pv)
            modified = macaulay / (1 + y / freq)
            convexity = compute_convexity(risk_cf, risk_t, y, dirty_pv, freq, False)
        except ZeroDivisionError:
            macaulay = modified = convexity = 0.0
    
        return pv, macaulay, modified, convexity

    min_yield = max(0.0001, ytm - 0.025)
    max_yield = ytm + 0.025
    yields = np.linspace(min_yield, max_yield, 100)

    prices, convexities = [], []
    for y in yields:
        dirty, _, _, cx = bond_pv(y)
        prices.append(dirty)
        convexities.append(cx)

    x_vals = np.array([y * 100 for y in yields])
    y_conv = np.array(convexities)
    spline = make_interp_spline(x_vals, y_conv, k=3)
    x_smooth = np.linspace(x_vals.min(), x_vals.max(), 300)
    y_smooth = spline(x_smooth)

    fig = plt.Figure(figsize=(8, 4.5))
    ax = fig.add_subplot(111)
    ax2 = ax.twinx()

    ax.plot(x_vals, prices, color="blue", label="Dirty Price Curve")
    ax.axhline(total_amt, color='red', linestyle='--', label=f'Observed Dirty Price: ${total_amt:,.2f}')
    ax.axvline(ytm * 100, color='green', linestyle='--', label=f'Yield: {ytm * 100:.2f}%')
    ax.scatter(ytm * 100, total_amt, color='black', label="Current Price Point")

    ax2.plot(x_smooth, y_smooth, color="orange", label="Convexity Curve (Smoothed)")

    ax.set_title("Bond Dirty Price and Convexity vs Yield to Maturity")
    ax.set_xlabel("Yield to Maturity (%)")
    ax.set_ylabel("Dirty Price (per $ Face Value)", color="blue")
    ax2.set_ylabel("Convexity (Duration² Units)", color="orange")

    ax.legend(loc="upper left")
    ax2.legend(loc="upper right")
    ax.grid(True)

    premium_or_discount = total_amt - face_value
    info = (
        f"Settlement Value: ${total_amt:,.2f}\n"
        f"Accrued Interest: ${accrued:,.2f}\n"
        f"{'Premium' if premium_or_discount > 0 else 'Discount'}: ${abs(premium_or_discount):,.2f}\n\n"
        f"Duration: {dura:.3f}\n"
        f"Modified Duration: {mdura:.3f}\n"
        f"Convexity: {conv:.3f}"
    )
    fig.text(0.72, 0.4, info, fontsize=10, bbox=dict(boxstyle="round", facecolor="white", alpha=0.9))

    popup = tk.Toplevel(root)
    popup.title("Bond Price and Convexity Chart")
    popup.geometry("850x550")
    popup.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - (popup.winfo_width() // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (popup.winfo_height() // 2)
    popup.geometry(f"+{x}+{y}")

    canvas = FigureCanvasTkAgg(fig, master=popup)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)
    canvas.get_tk_widget().bind("<Enter>", lambda e: canvas.get_tk_widget().focus_force())
    canvas.get_tk_widget().configure(takefocus=True)
    
    # ----- Annotations -----
  
    lock_annotation = ax.annotate(
        "", xy=(0, 0), xytext=(15, 15), textcoords="offset points",
        bbox=dict(boxstyle="round", fc="lightblue", alpha=0.9),
        arrowprops=dict(arrowstyle="->"), fontsize=9
    )
    lock_annotation.set_visible(False)

    locked = [False]
    lock_index = [-1]
          

    def on_click(event):
            
        if event.inaxes not in [ax, ax2] or event.x is None or event.y is None:
            # Click outside axes — clear label
            lock_annotation.set_visible(False)
            locked[0] = False
            lock_index[0] = -1
            canvas.draw_idle()
            return
    
        # 1️⃣ Convert the clicked location into pixel coordinates
        click_xy = np.array([event.x, event.y])
    
        # 2️⃣ Convert data points from (x_vals, prices) to pixel space using ax.transData
        pixel_points = np.array([
            ax.transData.transform((x, y)) for x, y in zip(x_vals, prices)
        ])
    
        # 3️⃣ Compute distances in pixel space
        distances = np.linalg.norm(pixel_points - click_xy, axis=1)
        idx = np.argmin(distances)
        
    
        # 4️⃣ Filter: Only lock if click is within a threshold distance (e.g., 15 pixels)
        if distances[idx] > 15:
            lock_annotation.set_visible(False)
            locked[0] = False
            lock_index[0] = -1
            canvas.draw_idle()
            return
    
        #  Toggle annotation if same point clicked again
        if locked[0] and idx == lock_index[0]:
            lock_annotation.set_visible(False)
            locked[0] = False
            lock_index[0] = -1
            canvas.draw_idle()
            return
    
        # Otherwise, display and lock annotation at nearest point
        x_val, y_val = x_vals[idx], prices[idx]
        lock_annotation.xy = (x_val, y_val)
        lock_annotation.set_text(f"Yield: {x_val:.2f}%\nPrice: ${y_val:,.2f}")
        lock_annotation.set_visible(True)
        locked[0] = True
        lock_index[0] = idx
        canvas.draw_idle()
        
        
    canvas.mpl_connect("button_press_event", on_click)

    canvas.get_tk_widget().bind("<Enter>", lambda e: canvas.get_tk_widget().focus_set())
    def save_chart_as_png():
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")],
            title="Save Chart as PNG"
        )
        if file_path:
            fig.savefig(file_path)
            print(f"Chart saved to {file_path}")

    tk.Button(popup, text="Save as PNG", command=save_chart_as_png).pack(pady=5)

def calculate_then_generate_chart():
    calculate()
    generate_price_yield_chart()

generate_button.config(command=calculate_then_generate_chart)

class PDFWithFooter(FPDF):
    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", size=8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")

def preview_amortization_pdf():
    try:
        pdf = PDFWithFooter()
        pdf.add_page()
        pdf.set_font("Arial", size=10)

        # Header and Inputs
        pdf.set_font("Arial", style="B", size=14)
        pdf.cell(0, 10, txt="Northern Territory Treasury Corporation", ln=True, align="C")
        pdf.cell(0, 10, txt="Bond Price Calculator Results", ln=True, align="C")
        pdf.ln(5)
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 10, txt=f"Printed on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True)

        pdf.ln(4)
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 10, txt="--- Input Data ---", ln=True)
        pdf.set_font("Arial", size=10)
        for i, label in enumerate(["Settlement Date", "Maturity Date", "Yield", "Coupon Rate", "Face Value"]):
            val = entry_vars[i].get()
            pdf.cell(0, 10, txt=f"{label}: {val}", ln=True)

        pdf.cell(0, 10, txt=f"Coupon Frequency: {coupon_frequency_var.get()}", ln=True)
        pdf.cell(0, 10, txt=f"Day Count Convention: {day_count_var.get()}", ln=True)
        pdf.cell(0, 10, txt=f"Interest Status: {interest_status_var.get()}", ln=True)
        pdf.cell(0, 10, txt=f"Ex-Interest Period (days): {ex_interest_days_var.get()}", ln=True)

        pdf.ln(4)
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 10, txt="--- Price per $100 ---", ln=True)
        pdf.set_font("Arial", size=10)
        for i, label in enumerate(["Capital Price", "Accrued Interest", "Gross Price"]):
            val = price_labels[i].cget("text")
            pdf.cell(0, 10, txt=f"{label}: {val}", ln=True)

        pdf.ln(4)
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 10, txt="--- Settlement Values ---", ln=True)
        pdf.set_font("Arial", size=10)
        for i, label in enumerate(["Capital Amount", "Accrued Interest Amount", "Total Settlement Value"]):
            val = settlement_labels[i].cget("text")
            pdf.cell(0, 10, txt=f"{label}: {val}", ln=True)

        pdf.ln(4)
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 10, txt="--- Risk Metrics ---", ln=True)
        pdf.set_font("Arial", size=10)
        for i, label in enumerate(["Duration", "Modified Duration", "Convexity"]):
            val = risk_labels[i].cget("text")
            pdf.cell(0, 10, txt=f"{label}: {val}", ln=True)

        # Amortization Schedule
        pdf.ln(4)
        pdf.set_font("Arial", style="B", size=12)
        pdf.cell(0, 10, txt="--- Amortization Schedule ---", ln=True)
        pdf.set_font("Courier", size=8)

        lines = amort_lines.copy()
        lines_per_page = 48
        for i, line in enumerate(lines):
            if i > 0 and i % lines_per_page == 0:
                pdf.add_page()
                pdf.set_font("Courier", size=8)
            pdf.cell(0, 5, txt=line.strip(), ln=True)

        # Save to a temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            temp_path = tmp.name
            pdf.output(temp_path)

        # Open with system default PDF viewer
        if platform.system() == "Darwin":
            os.system(f"open '{temp_path}'")
        elif platform.system() == "Windows":
            os.startfile(temp_path)
        else:
            os.system(f"xdg-open '{temp_path}'")

    except Exception as e:
        print(f"Error previewing PDF: {e}")

def calculate_then_preview_pdf():
    calculate()
    preview_amortization_pdf()

export_button.config(command=calculate_then_preview_pdf)

def global_click(event):
    normalize_currency_field(entry_vars[4])
    normalize_rate_field(entry_vars[2])
    normalize_rate_field(entry_vars[3])

def enable_auto_calculation():
    global auto_calc_enabled
    if not auto_calc_enabled:
        auto_calc_enabled = True
        attach_auto_calculate()

def calculate_then_enable_auto():
    calculate(manual=True)
    enable_auto_calculation()

def calculate_then_generate_chart():
    calculate(manual=True)
    generate_price_yield_chart()

def calculate_then_preview_pdf():
    calculate(manual=True)
    preview_amortization_pdf()

calculate_button.config(command=calculate_then_enable_auto)

root.bind_all("<Button-1>", global_click, add="+")


# amort_lines = [f"Test Row {i}" for i in range(24)]
# display_amort_page(0)

if __name__ == "__main__":
    root.mainloop()
