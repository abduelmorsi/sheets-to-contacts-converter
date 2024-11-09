import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import requests

class ContactConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Contact Converter")
        self.root.minsize(400, 300)  
        
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        
        ttk.Label(main_frame, text="CSV File Path or Published Google Sheet URL:").grid(row=0, column=0, sticky="w")
        self.sheet_url = ttk.Entry(main_frame)
        self.sheet_url.grid(row=1, column=0, sticky="ew", pady=2)
        ttk.Button(main_frame, text="Browse", command=self.browse_file).grid(row=2, column=0, pady=2)
        ttk.Label(main_frame, text="Name Prefix (optional):").grid(row=3, column=0, sticky="w")
        self.name_prefix = ttk.Entry(main_frame)
        self.name_prefix.grid(row=4, column=0, sticky="ew", pady=2)

        column_frame = ttk.LabelFrame(main_frame, text="Column Selection", padding=10)
        column_frame.grid(row=5, column=0, sticky="ew", pady=5)
        column_frame.grid_columnconfigure(0, weight=1)

        ttk.Label(column_frame, text="Name Column Letter:").grid(row=0, column=0, sticky="w")
        self.name_col = ttk.Entry(column_frame, width=5)
        self.name_col.grid(row=1, column=0, sticky="w", pady=2)
        self.name_col.insert(0, "B")

        ttk.Label(column_frame, text="Phone Column Letter:").grid(row=2, column=0, sticky="w")
        self.phone_col = ttk.Entry(column_frame, width=5)
        self.phone_col.grid(row=3, column=0, sticky="w", pady=2)
        self.phone_col.insert(0, "C")

        format_frame = ttk.LabelFrame(main_frame, text="Output Format", padding=10)
        format_frame.grid(row=6, column=0, sticky="ew", pady=5)
        self.format_var = tk.StringVar(value="csv")
        ttk.Radiobutton(format_frame, text="CSV", variable=self.format_var, value="csv").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(format_frame, text="VCF", variable=self.format_var, value="vcf").grid(row=1, column=0, sticky="w")

        ttk.Button(main_frame, text="Convert", command=self.convert_contacts).grid(row=7, column=0, pady=10)

        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=8, column=0, sticky="ew")

    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.sheet_url.delete(0, tk.END)
            self.sheet_url.insert(0, filename)

    def convert_contacts(self):
        try:
            sheet_url = self.sheet_url.get().strip()
            if not sheet_url:
                raise ValueError("File path or URL is required")
            if not self.name_col.get().strip() or not self.phone_col.get().strip():
                raise ValueError("Column letters are required")

            if "docs.google.com/spreadsheets" in sheet_url:
                sheet_id = sheet_url.split('/d/')[1].split('/')[0]
                csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                response = requests.get(csv_url)
                if response.status_code != 200:
                    raise ValueError("Unable to access the sheet. Make sure it's published to the web.")
                df = pd.read_csv(pd.io.common.StringIO(response.text))
            else:
                df = pd.read_csv(sheet_url)

            name_col_letter = self.name_col.get().upper()
            phone_col_letter = self.phone_col.get().upper()
            
            columns = df.columns.tolist()
            name_col_idx = ord(name_col_letter) - ord('A')
            phone_col_idx = ord(phone_col_letter) - ord('A')
            
            if name_col_idx >= len(columns) or phone_col_idx >= len(columns):
                raise ValueError("Column letter is out of range")
                
            name_col = columns[name_col_idx]
            phone_col = columns[phone_col_idx]

            prefix = self.name_prefix.get()
            if prefix:
                df[name_col] = prefix + " " + df[name_col]

            if self.format_var.get() == "csv":
                self.save_as_csv(df, name_col, phone_col)
            else:
                self.save_as_vcf(df, name_col, phone_col)

            messagebox.showinfo("Success", "Contacts file created successfully!")
            
        except FileNotFoundError:
            messagebox.showerror("Error", "CSV file not found")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def save_as_csv(self, df, name_col, phone_col):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            output = pd.DataFrame({
                'Name': df[name_col],
                'Given Name': '',
                'Additional Name': '',
                'Family Name': '',
                'Mobile Phone': df[phone_col]
            })
            output.to_csv(file_path, index=False, encoding='utf-8-sig')

    def save_as_vcf(self, df, name_col, phone_col):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".vcf",
            filetypes=[("VCF files", "*.vcf")]
        )
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                for _, row in df.iterrows():
                    f.write('BEGIN:VCARD\n')
                    f.write('VERSION:3.0\n')
                    f.write(f'FN:{row[name_col]}\n')
                    f.write(f'TEL;TYPE=CELL:{row[phone_col]}\n')
                    f.write('END:VCARD\n\n')

if __name__ == "__main__":
    root = tk.Tk()
    app = ContactConverterApp(root)
    root.mainloop()
