import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, Button, filedialog, Label
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage

df = None  

def load_csv():
    global df
    file_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV Files", "*.csv")])
    if file_path:
        df = pd.read_csv(file_path)
        status_label.config(text=f"Loaded: {file_path.split('/')[-1]} ✅")

def generate_report():
    global df
    if df is None:
        status_label.config(text="⚠️ Please load a CSV first!")
        return

    pivot = pd.pivot_table(df, index="Region", columns="Product", values="Sales", aggfunc="sum", fill_value=0)

    plt.figure(figsize=(6,4))
    pivot.plot(kind="bar", stacked=True)
    plt.title("Sales by Region & Product")
    plt.ylabel("Total Sales")
    plt.tight_layout()
    chart_file = "chart.png"
    plt.savefig(chart_file)
    plt.close()

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Pivot Table"

    for r in dataframe_to_rows(pivot, index=True, header=True):
        ws1.append(r)

    img = XLImage(chart_file)
    ws1.add_image(img, "H2")

    ws2 = wb.create_sheet("Summary")
    total_sales = df["Sales"].sum()
    avg_sales = df["Sales"].mean()
    top_product = df.groupby("Product")["Sales"].sum().idxmax()
    ws2.append(["Total Sales", total_sales])
    ws2.append(["Average Sales", avg_sales])
    ws2.append(["Top Product", top_product])

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")])
    if save_path:
        wb.save(save_path)
        status_label.config(text=f"Report saved: {save_path.split('/')[-1]} ✅")

root = Tk()
root.title("Excel Report Generator")
root.geometry("400x200")

Label(root, text="📊 Excel Report Generator", font=("Arial", 14, "bold")).pack(pady=10)

Button(root, text="📂 Load CSV", command=load_csv, width=20, bg="lightblue").pack(pady=5)
Button(root, text="📑 Generate Report", command=generate_report, width=20, bg="lightgreen").pack(pady=5)

status_label = Label(root, text="Please load a CSV file", fg="gray")
status_label.pack(pady=10)

root.mainloop()