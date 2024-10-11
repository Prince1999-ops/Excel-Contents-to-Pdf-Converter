import os
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from fpdf import FPDF

file_path = r"C:\Users\pmafuru.TTPL\Desktop\Book1tests.xlsx"  
try:
    data = pd.read_excel(file_path)
except FileNotFoundError:
    print("The specified file was not found. Please check the file path.")
    exit()

data.columns = data.columns.str.strip()
if 'PS' not in data.columns:
    raise ValueError("The 'PS' column is not found in the Excel file. Please check the data.")

unique_ps = pd.Series(data['PS'].unique()).str.strip().tolist()

def save_ps_to_pdf():
    selected_ps = ps_combobox.get()
    if selected_ps:
        ps_info = data[data['PS'] == selected_ps]
        print(f"Information for {selected_ps}:")
        print(ps_info)

        output_folder_path = filedialog.askdirectory(title="Select Folder to Save PDF")
        if output_folder_path:
            output_file_path = os.path.join(output_folder_path, f"{selected_ps}_information.pdf")
            save_to_pdf(ps_info, output_file_path)
            messagebox.showinfo("Success", f"Information saved to {output_file_path}")
        else:
            messagebox.showwarning("File Save Error", "No folder selected.")
    else:
        messagebox.showwarning("Selection Error", "Please select a Primary Society.")

def wrap_text(text, col_width, pdf):
    wrapped_lines = []
    words = text.split(' ')
    line = ""

    for word in words:
        if pdf.get_string_width(line + word + ' ') < col_width:
            line += word + ' '
        else:
            wrapped_lines.append(line.strip())
            line = word + ' '
    wrapped_lines.append(line.strip())

    return wrapped_lines

def save_to_pdf(df, output_file_path):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Arial", "B", 12)
    pdf.cell(200, 10, f"Barn Information for {df['PS'].iloc[0]}", ln=True, align='C')
    pdf.ln(10)

    pdf.set_font("Arial", "B", 7)
    col_widths = []
    table_width = pdf.w - 20
    col_max_width = table_width / len(df.columns)
    
    for col in df.columns:
        max_width = max(df[col].astype(str).map(len).max(), len(col))
        col_width = max(10, min(col_max_width, max_width * 1.5))
        col_widths.append(col_width)

    for i, column in enumerate(df.columns):
        pdf.cell(col_widths[i], 10, column, border=1, align='C')
    pdf.ln()

    pdf.set_font("Arial", "", 7)
    row_height = pdf.font_size * 2
    for row in df.itertuples(index=False):
        max_lines = 1
        cell_data_wrapped = []

        for i, value in enumerate(row):
            if pd.isna(value):
                wrapped_lines = wrap_text('NaN', col_widths[i], pdf)
            elif isinstance(value, float):
                value = int(value)
                wrapped_lines = wrap_text(str(value), col_widths[i], pdf)
            else:
                wrapped_lines = wrap_text(str(value), col_widths[i], pdf)

            cell_data_wrapped.append(wrapped_lines)
            max_lines = max(max_lines, len(wrapped_lines))

        for line_num in range(max_lines):
            for i, wrapped_lines in enumerate(cell_data_wrapped):
                if line_num < len(wrapped_lines):
                    pdf.cell(col_widths[i], row_height, wrapped_lines[line_num], border=1)
                else:
                    pdf.cell(col_widths[i], row_height, '', border=1)
            pdf.ln(row_height)

    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)

    try:
        pdf.output(output_file_path)
    except PermissionError:
        messagebox.showerror("Permission Error", f"Unable to save PDF to {output_file_path}. Please check your permissions.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

root = tk.Tk()
root.title("Select Primary Society")

label = tk.Label(root, text="Select Primary Society:")
label.pack(pady=10)

ps_combobox = ttk.Combobox(root, values=unique_ps)
ps_combobox.pack(pady=10)

save_button = tk.Button(root, text="Save Information as PDF", command=save_ps_to_pdf)
save_button.pack(pady=20)

root.mainloop()
