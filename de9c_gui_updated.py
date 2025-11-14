import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject
import os
import math
import datetime
import json

ROWS_PER_PAGE = 7
CONFIG_FILE = "de9c_defaults.json"


# --------------------------------------------------------
# Utility functions
# --------------------------------------------------------

def clean_money(x):
    return str(x).replace("$", "").replace(",", "").strip()


def suffix_for_row(i: int) -> str:
    """Row 1 -> '', Row 2 -> '1', Row 3 -> '2', ..."""
    return "" if i == 1 else str(i - 1)


def calc_quarter_end(year_full: int, q: int) -> str:
    yy = str(year_full)[-2:]
    if q == 1:
        return f"03/31/{yy}"
    if q == 2:
        return f"06/30/{yy}"
    if q == 3:
        return f"09/30/{yy}"
    if q == 4:
        return f"12/31/{yy}"
    raise ValueError("Quarter must be 1–4")


def quarter_months(q: int):
    """Return (m1, m2, m3) month numbers for a quarter."""
    mapping = {
        1: (1, 2, 3),
        2: (4, 5, 6),
        3: (7, 8, 9),
        4: (10, 11, 12),
    }
    if q not in mapping:
        raise ValueError("Quarter must be 1–4")
    return mapping[q]


# --------------------------------------------------------
# Config (JSON) helpers
# --------------------------------------------------------

def load_defaults():
    """Load defaults from JSON if present, otherwise return hard-coded defaults."""
    defaults = {
        "year": "2024",
        "quarter": "2",
        "employer_account": "14503726",
        "signature_name": "EDUARDO b PEREZ",
        "signature_title": "PRESIDENT",
        "signature_phone": "818-618-1851",
    }
    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            defaults.update(data)
        except Exception:
            # If anything goes wrong, fall back to built-in defaults
            pass
    return defaults


def save_defaults(year, quarter, account, sig_name, sig_title, sig_phone):
    data = {
        "year": str(year),
        "quarter": str(quarter),
        "employer_account": account,
        "signature_name": sig_name,
        "signature_title": sig_title,
        "signature_phone": sig_phone,
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print("Warning: could not save defaults:", e)


# --------------------------------------------------------
# Main DE-9C filling logic
# --------------------------------------------------------

def fill_de9c(csv_path, template_path, output_path,
              year_full, quarter, acct, quarter_end,
              sig_name, sig_title, sig_phone, sig_date):

    # Load CSV
    df = pd.read_csv(csv_path)

    # Clean money fields
    for col in ["Total Subject Wages", "PIT Wages", "PIT Withheld"]:
        df[col] = df[col].astype(str).map(clean_money)

    # Build employee list
    employees = []
    for _, row in df.iterrows():
        middle_raw = row.get("Middle Name", "")
        if pd.isna(middle_raw):
            middle_value = ""
        else:
            middle_value = str(middle_raw).strip()

        employees.append(
            [
                row["SSN"],
                row["First Name"],
                middle_value,
                row["Last Name"],
                row["Total Subject Wages"],
                row["PIT Wages"],
                row["PIT Withheld"],
            ]
        )

    num_employees = len(employees)
    num_pages = math.ceil(num_employees / ROWS_PER_PAGE)

    FIELD_MAP = {
        "SSN": "SSN{suffix}",
        "First": "First Name{suffix}",
        "MI": "MI{suffix}",
        "Last": "Last Name{suffix}",
        "Wages": "Total Subject Wages{suffix}",
        "PITWages": "PIT Wages{suffix}",
        "Withheld": "PIT Withheld{suffix}",
    }

    filled_pages = []
    grand_w = grand_p = grand_h = 0.0

    # Month boxes based on quarter
    m1, m2, m3 = quarter_months(quarter)
    m1_str = f"{m1:02d}"
    m2_str = f"{m2:02d}"
    m3_str = f"{m3:02d}"

    for page_idx in range(num_pages):
        start = page_idx * ROWS_PER_PAGE
        page_emps = employees[start:start + ROWS_PER_PAGE]

        reader = PdfReader(template_path)
        page = reader.pages[0]
        annots = page.get("/Annots") or []
        lookup = {a["/T"][1:-1]: a for a in annots if a.get("/T")}

        # -------- Header fields --------
        year2 = str(year_full)[-2:]

        if "Year" in lookup:
            lookup["Year"].update(PdfDict(V=year2, AS=year2, AP=None))

        if "Quarter" in lookup:
            lookup["Quarter"].update(PdfDict(V=str(quarter), AS=str(quarter), AP=None))

        if "Employer Account No" in lookup:
            lookup["Employer Account No"].update(
                PdfDict(V=acct, AS=acct, AP=None)
            )

        # Quarter Ended (Date1)
        if "Date1" in lookup:
            lookup["Date1"].update(PdfDict(V=quarter_end, AS=quarter_end, AP=None))

        # DUE (Date2) – same as quarter ended
        if "Date2" in lookup:
            lookup["Date2"].update(PdfDict(V=quarter_end, AS=quarter_end, AP=None))

        # Month boxes
        if "1st Month" in lookup:
            lookup["1st Month"].update(PdfDict(V=m1_str, AS=m1_str, AP=None))
        if "2nd Month" in lookup:
            lookup["2nd Month"].update(PdfDict(V=m2_str, AS=m2_str, AP=None))
        if "3rd Month" in lookup:
            lookup["3rd Month"].update(PdfDict(V=m3_str, AS=m3_str, AP=None))

        # Page numbering
        pno = str(page_idx + 1)
        tpages = str(num_pages)
        if "Page number" in lookup:
            lookup["Page number"].update(PdfDict(V=pno, AS=pno, AP=None))
        if "Of Page number" in lookup:
            lookup["Of Page number"].update(PdfDict(V=tpages, AS=tpages, AP=None))

        # -------- Rows --------
        page_w = page_p = page_h = 0.0

        for r, emp in enumerate(page_emps, start=1):
            suf = suffix_for_row(r)

            data = {
                "SSN": emp[0],
                "First": emp[1],
                "MI": emp[2],
                "Last": emp[3],
                "Wages": emp[4],
                "PITWages": emp[5],
                "Withheld": emp[6],
            }

            try:
                page_w += float(emp[4] or "0")
            except Exception:
                pass

            try:
                page_p += float(emp[5] or "0")
            except Exception:
                pass

            try:
                page_h += float(emp[6] or "0")
            except Exception:
                pass

            for key, val in data.items():
                fname = FIELD_MAP[key].format(suffix=suf)
                if fname in lookup:
                    lookup[fname].update(PdfDict(V=str(val), AS=str(val), AP=None))

        # Clear unused rows on last page
        for blank_row in range(len(page_emps) + 1, ROWS_PER_PAGE + 1):
            suf = suffix_for_row(blank_row)
            for patt in FIELD_MAP.values():
                fname = patt.format(suffix=suf)
                if fname in lookup:
                    lookup[fname].update(PdfDict(V="", AS="", AP=None))

        # Rename page field names (except the first page) so cloned pages
        # don't share the same field names/form values.
        if page_idx > 0:
            suffix = f"__p{page_idx + 1}"
            for annot in annots:
                field_name = annot.get("/T")
                if not field_name:
                    continue
                original = field_name[1:-1]
                annot.update(PdfDict(T=PdfObject(f"({original}{suffix})")))

        # -------- Page totals (I/J/K) --------
        if "Total Subject Wages This Page" in lookup:
            txt = f"{page_w:.2f}"
            lookup["Total Subject Wages This Page"].update(
                PdfDict(V=txt, AS=txt, AP=None)
            )
        if "Total PIT Wages This Page" in lookup:
            txt = f"{page_p:.2f}"
            lookup["Total PIT Wages This Page"].update(
                PdfDict(V=txt, AS=txt, AP=None)
            )
        if "Total PIT Withheld This Page" in lookup:
            txt = f"{page_h:.2f}"
            lookup["Total PIT Withheld This Page"].update(
                PdfDict(V=txt, AS=txt, AP=None)
            )

        grand_w += page_w
        grand_p += page_p
        grand_h += page_h

        filled_pages.append(reader)

    # -------- Grand totals & signature on Page 1 --------
    first_reader = filled_pages[0]
    first_page = first_reader.pages[0]
    annots = first_page.get("/Annots") or []
    lookup = {a["/T"][1:-1]: a for a in annots if a.get("/T")}

    # Grand totals L/M/N
    if "Grand Total Subject Wages" in lookup:
        txt = f"{grand_w:.2f}"
        lookup["Grand Total Subject Wages"].update(
            PdfDict(V=txt, AS=txt, AP=None)
        )
    if "Grand Total PIT Wages" in lookup:
        txt = f"{grand_p:.2f}"
        lookup["Grand Total PIT Wages"].update(
            PdfDict(V=txt, AS=txt, AP=None)
        )
    if "Grand Total PIT Withheld" in lookup:
        txt = f"{grand_h:.2f}"
        lookup["Grand Total PIT Withheld"].update(
            PdfDict(V=txt, AS=txt, AP=None)
        )

    # Signature block – your PDF uses these internal names:
    # Signature1 = Signature
    # Text2      = Title
    # Text3      = Phone
    # 0          = Date
    sig_map = {
        "Signature1": sig_name,
        "0": sig_title,                  # Title field
        "Phone Number": sig_phone,       # Phone
        "Date5": sig_date,               # Date
    }

    for fname, val in sig_map.items():
        if fname in lookup:
            lookup[fname].update(PdfDict(V=val, AS=val, AP=None))

    # Ensure appearances are regenerated
    if hasattr(first_reader, "Root") and first_reader.Root.AcroForm:
        first_reader.Root.AcroForm.update(
            PdfDict(NeedAppearances=PdfObject("true"))
        )

    # -------- Merge & write --------
    writer = PdfWriter()
    for r in filled_pages:
        writer.addpage(r.pages[0])

    final_trailer = writer.trailer
    if hasattr(first_reader, "Root") and first_reader.Root.AcroForm:
        final_trailer.Root.AcroForm = first_reader.Root.AcroForm

    writer.write(output_path, trailer=final_trailer)


# --------------------------------------------------------
# GUI
# --------------------------------------------------------

def run_gui():
    root = tk.Tk()
    root.title("DE-9C Autofill Tool")
    root.geometry("680x760")

    defaults = load_defaults()

    # Variables
    csv_var = tk.StringVar()
    pdf_var = tk.StringVar()
    out_var = tk.StringVar()

    year_var = tk.StringVar(value=defaults["year"])
    qtr_var = tk.StringVar(value=defaults["quarter"])
    acct_var = tk.StringVar(value=defaults["employer_account"])
    qend_var = tk.StringVar(value="")  # if blank, auto-calc

    name_var = tk.StringVar(value=defaults["signature_name"])
    title_var = tk.StringVar(value=defaults["signature_title"])
    phone_var = tk.StringVar(value=defaults["signature_phone"])
    today = datetime.date.today().strftime("%m/%d/%y")
    date_var = tk.StringVar(value=today)

    # Browse helpers
    def browse_csv():
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            csv_var.set(path)

    def browse_pdf():
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            pdf_var.set(path)

    def browse_out():
        path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if path:
            out_var.set(path)

    def run_fill():
        try:
            if not csv_var.get():
                messagebox.showerror("Error", "Please select a CSV file.")
                return
            if not pdf_var.get():
                messagebox.showerror("Error", "Please select a DE-9C template PDF.")
                return
            if not out_var.get():
                messagebox.showerror("Error", "Please choose an output PDF file.")
                return

            year_full = int(year_var.get())
            q = int(qtr_var.get())
            if q not in (1, 2, 3, 4):
                raise ValueError("Quarter must be 1–4.")

            quarter_end = qend_var.get().strip() or calc_quarter_end(year_full, q)

            fill_de9c(
                csv_var.get(),
                pdf_var.get(),
                out_var.get(),
                year_full,
                q,
                acct_var.get(),
                quarter_end,
                name_var.get(),
                title_var.get(),
                phone_var.get(),
                date_var.get(),
            )

            # Save current values as defaults
            save_defaults(
                year_var.get(),
                qtr_var.get(),
                acct_var.get(),
                name_var.get(),
                title_var.get(),
                phone_var.get(),
            )

            messagebox.showinfo("Success", "DE-9C PDF created successfully!")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # Layout
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill="both", expand=True)

    row = 0
    for label, var, cmd in [
        ("CSV File:", csv_var, browse_csv),
        ("DE-9C Template PDF:", pdf_var, browse_pdf),
        ("Output PDF:", out_var, browse_out),
    ]:
        tk.Label(frame, text=label).grid(row=row, column=0, sticky="w")
        tk.Entry(frame, textvariable=var, width=40).grid(row=row, column=1)
        tk.Button(frame, text="Browse", command=cmd).grid(row=row, column=2)
        row += 1

    # Input fields
    for label, var in [
        ("Tax Year:", year_var),
        ("Quarter (1-4):", qtr_var),
        ("Employer Account Number:", acct_var),
        ("Quarter Ended (mm/dd/yy, blank = auto):", qend_var),
        ("Signature Name:", name_var),
        ("Signature Title:", title_var),
        ("Signature Phone:", phone_var),
        ("Signature Date:", date_var),
    ]:
        tk.Label(frame, text=label).grid(row=row, column=0, sticky="w")
        tk.Entry(frame, textvariable=var, width=40).grid(
            row=row, column=1, columnspan=2
        )
        row += 1

    tk.Button(
        frame,
        text="GENERATE DE-9C",
        command=run_fill,
        bg="#4CAF50",
        fg="white",
        height=2,
    ).grid(row=row, column=0, columnspan=3, pady=20)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
