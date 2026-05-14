import os
import pandas as pd
import qrcode
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import tkinter as tk
from tkinter import filedialog, messagebox


def col_letter(n):
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def generate_qr_excel():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, header=None)

        output_file = os.path.join(
            os.path.dirname(file_path),
            "QR_Output.xlsx"
        )

        wb = Workbook()
        ws = wb.active

        temp_folder = "qr_temp"
        os.makedirs(temp_folder, exist_ok=True)

        for row in range(df.shape[0]):
            for col in range(df.shape[1]):

                data = str(df.iat[row, col])

                qr = qrcode.make(data)

                img_path = os.path.join(
                    temp_folder,
                    f"qr_{row}_{col}.png"
                )

                qr.save(img_path)

                img = Image(img_path)
                img.width = 80
                img.height = 80

                cell = ws.cell(
                    row=row + 1,
                    column=col + 1
                ).coordinate

                ws.add_image(img, cell)

                ws.row_dimensions[row + 1].height = 65
                ws.column_dimensions[
                    col_letter(col + 1)
                ].width = 15

        wb.save(output_file)

        messagebox.showinfo(
            "Success",
            f"QR Excel Created:\n{output_file}"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI
root = tk.Tk()
root.title("Excel QR Generator")
root.geometry("400x200")

label = tk.Label(
    root,
    text="Excel to QR Code Generator",
    font=("Arial", 16)
)
label.pack(pady=20)

btn = tk.Button(
    root,
    text="Select Excel File",
    command=generate_qr_excel,
    width=25,
    height=2
)
btn.pack(pady=20)

root.mainloop()