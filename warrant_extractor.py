"""
Property Alias Warrant Data Extractor
Extracts Name, Address, City/State/Zip, and Total from warrant PDFs
and outputs them to an Excel spreadsheet.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber is required. Install with: pip install pdfplumber")
    sys.exit(1)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except ImportError:
    print("ERROR: openpyxl is required. Install with: pip install openpyxl")
    sys.exit(1)


def extract_warrant_data(pdf_path):
    """
    Extract Name, Address, City/State/Zip, and Total from each page of a warrant PDF.
    
    Uses two strategies:
    1. Position-based extraction (primary) - uses known coordinates from the form layout
    2. Text-pattern extraction (fallback) - parses the text content looking for known patterns
    
    Returns a list of dicts, one per page.
    """
    results = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            record = {
                "source_file": os.path.basename(pdf_path),
                "page": page_num,
                "name": "",
                "address": "",
                "city": "",
                "state": "",
                "zip_code": "",
                "total": "",
            }

            words = page.extract_words()

            if words:
                # --- Strategy 1: Position-based extraction ---
                # The form layout is consistent:
                #   Name line:       top ~130-136, x0 >= 70
                #   Address line:    top ~145-150, x0 >= 70
                #   City/State/Zip:  top ~160-165, x0 >= 70
                #   Total amount:    top ~455-462, x0 >= 470

                name_words = [w for w in words if 125 < w["top"] < 140 and w["x0"] >= 65]
                addr_words = [w for w in words if 140 < w["top"] < 155 and w["x0"] >= 65]
                csz_words = [w for w in words if 155 < w["top"] < 170 and w["x0"] >= 65]
                total_words = [w for w in words if 450 < w["top"] < 470 and w["x0"] >= 460]

                record["name"] = " ".join(w["text"] for w in name_words).strip()
                record["address"] = " ".join(w["text"] for w in addr_words).strip()
                csz_full = " ".join(w["text"] for w in csz_words).strip()
                # Parse "MANCHESTER CT 06040-1234" into city, state, zip
                csz_parts = csz_full.rsplit(" ", 2)  # split from right: [city, state, zip]
                if len(csz_parts) == 3:
                    record["city"] = csz_parts[0]
                    record["state"] = csz_parts[1]
                    record["zip_code"] = csz_parts[2]
                else:
                    record["city"] = csz_full

                if total_words:
                    # Pick the word that looks like a dollar amount
                    for tw in total_words:
                        text = tw["text"].replace(",", "").replace("$", "")
                        try:
                            float(text)
                            record["total"] = tw["text"]
                            break
                        except ValueError:
                            continue

            # --- Strategy 2: Fallback text-based extraction ---
            if not record["name"] or not record["total"]:
                text = page.extract_text() or ""
                lines = [l.strip() for l in text.split("\n") if l.strip()]

                # Look for the "collect forthwith from" marker
                for i, line in enumerate(lines):
                    if "collect forthwith from" in line.lower():
                        # Next 3 non-empty lines should be name, address, city/state/zip
                        remaining = [l for l in lines[i + 1 :] if l.strip()]
                        if len(remaining) >= 3 and not record["name"]:
                            record["name"] = remaining[0]
                            record["address"] = remaining[1]
                            csz_full = remaining[2]
                            csz_parts = csz_full.rsplit(" ", 2)
                            if len(csz_parts) == 3:
                                record["city"] = csz_parts[0]
                                record["state"] = csz_parts[1]
                                record["zip_code"] = csz_parts[2]
                            else:
                                record["city"] = csz_full
                        break

                # Look for total near "BALANCE DUE" line
                if not record["total"]:
                    for i, line in enumerate(lines):
                        if "balance due" in line.lower():
                            # Check surrounding lines for a dollar amount
                            for nearby in lines[max(0, i - 2) : i + 3]:
                                cleaned = nearby.replace(",", "").replace("$", "").strip()
                                try:
                                    float(cleaned)
                                    record["total"] = nearby.strip()
                                    break
                                except ValueError:
                                    continue
                            break

            results.append(record)

    return results


def write_to_excel(all_records, output_path):
    """Write extracted records to a formatted Excel spreadsheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Warrant Data"

    # Header style
    header_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Headers
    headers = ["Name", "Address", "City", "State", "Zip Code", "Total", "Source File", "Page"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # Data rows
    data_font = Font(name="Calibri", size=11)
    for row_idx, record in enumerate(all_records, start=2):
        values = [
            record["name"],
            record["address"],
            record["city"],
            record["state"],
            record["zip_code"],
            record["total"],
            record["source_file"],
            record["page"],
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.border = thin_border

    # Auto-fit column widths
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_length + 4, 50)

    wb.save(output_path)


class WarrantExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Warrant Data Extractor")
        self.root.geometry("700x500")
        self.root.resizable(True, True)

        # Set minimum size
        self.root.minsize(600, 400)

        self.pdf_files = []
        self.build_ui()

    def build_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Property Alias Warrant Extractor",
            font=("Calibri", 16, "bold"),
        )
        title_label.pack(pady=(0, 5))

        subtitle_label = ttk.Label(
            main_frame,
            text="Extract Name, Address, City/State/Zip, and Total from warrant PDFs",
            font=("Calibri", 10),
        )
        subtitle_label.pack(pady=(0, 15))

        # Buttons frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.select_btn = ttk.Button(
            btn_frame, text="Select PDF Files", command=self.select_files
        )
        self.select_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.folder_btn = ttk.Button(
            btn_frame, text="Select Folder", command=self.select_folder
        )
        self.folder_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.clear_btn = ttk.Button(
            btn_frame, text="Clear List", command=self.clear_files
        )
        self.clear_btn.pack(side=tk.LEFT)

        # File list
        list_frame = ttk.LabelFrame(main_frame, text="PDF Files", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.file_listbox = tk.Listbox(list_frame, font=("Calibri", 10))
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Status and progress
        self.status_var = tk.StringVar(value="Select PDF files or a folder to begin.")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, font=("Calibri", 10))
        status_label.pack(fill=tk.X, pady=(0, 5))

        self.progress = ttk.Progressbar(main_frame, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 10))

        # Extract button
        self.extract_btn = ttk.Button(
            main_frame,
            text="Extract Data to Spreadsheet",
            command=self.run_extraction,
        )
        self.extract_btn.pack(pady=(0, 5))

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if files:
            for f in files:
                if f not in self.pdf_files:
                    self.pdf_files.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            self.status_var.set(f"{len(self.pdf_files)} file(s) selected.")

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing PDFs")
        if folder:
            count = 0
            for fname in sorted(os.listdir(folder)):
                if fname.lower().endswith(".pdf"):
                    full_path = os.path.join(folder, fname)
                    if full_path not in self.pdf_files:
                        self.pdf_files.append(full_path)
                        self.file_listbox.insert(tk.END, fname)
                        count += 1
            self.status_var.set(f"{len(self.pdf_files)} file(s) selected. ({count} added from folder)")

    def clear_files(self):
        self.pdf_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.progress["value"] = 0
        self.status_var.set("List cleared. Select PDF files or a folder to begin.")

    def run_extraction(self):
        if not self.pdf_files:
            messagebox.showwarning("No Files", "Please select PDF files first.")
            return

        # Ask where to save
        output_path = filedialog.asksaveasfilename(
            title="Save Spreadsheet As",
            defaultextension=".xlsx",
            initialfile=f"warrant_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        if not output_path:
            return

        # Disable buttons during extraction
        self.extract_btn.configure(state=tk.DISABLED)
        self.select_btn.configure(state=tk.DISABLED)
        self.folder_btn.configure(state=tk.DISABLED)

        # Run extraction in a thread so GUI stays responsive
        thread = threading.Thread(target=self.do_extraction, args=(output_path,), daemon=True)
        thread.start()

    def do_extraction(self, output_path):
        all_records = []
        total_files = len(self.pdf_files)

        for idx, pdf_path in enumerate(self.pdf_files):
            self.root.after(
                0,
                lambda i=idx, p=pdf_path: self.status_var.set(
                    f"Processing ({i+1}/{total_files}): {os.path.basename(p)}"
                ),
            )

            try:
                records = extract_warrant_data(pdf_path)
                all_records.extend(records)
            except Exception as e:
                self.root.after(
                    0,
                    lambda p=pdf_path, err=e: messagebox.showerror(
                        "Error", f"Error processing {os.path.basename(p)}:\n{err}"
                    ),
                )

            progress_val = ((idx + 1) / total_files) * 100
            self.root.after(0, lambda v=progress_val: self.progress.configure(value=v))

        # Write results
        if all_records:
            try:
                write_to_excel(all_records, output_path)
                self.root.after(
                    0,
                    lambda: self.status_var.set(
                        f"Done! Extracted {len(all_records)} records to {os.path.basename(output_path)}"
                    ),
                )
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Success",
                        f"Extracted {len(all_records)} records from {total_files} file(s).\n\nSaved to:\n{output_path}",
                    ),
                )
            except Exception as e:
                self.root.after(
                    0, lambda: messagebox.showerror("Error", f"Error saving spreadsheet:\n{e}")
                )
        else:
            self.root.after(
                0, lambda: messagebox.showwarning("No Data", "No records were extracted from the selected files.")
            )

        # Re-enable buttons
        self.root.after(0, lambda: self.extract_btn.configure(state=tk.NORMAL))
        self.root.after(0, lambda: self.select_btn.configure(state=tk.NORMAL))
        self.root.after(0, lambda: self.folder_btn.configure(state=tk.NORMAL))


def main():
    root = tk.Tk()
    app = WarrantExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()