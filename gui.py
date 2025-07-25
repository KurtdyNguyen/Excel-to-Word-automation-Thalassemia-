import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
from thalassemia import process_thalassemia_excel


CONFIG_PATH = ".thal_config.json"

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_config(data):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def select_excel():
    filepath = filedialog.askopenfilename(
        title="Chọn file Excel đầu vào",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if filepath:
        excel_path_var.set(filepath)
        config = load_config()
        config["last_excel_path"] = filepath
        save_config(config)

def select_output_folder():
    folder = filedialog.askdirectory(title="Chọn thư mục để lưu kết quả")
    if folder:
        output_dir_var.set(folder)
        config = load_config()
        config["last_output_dir"] = folder
        save_config(config)

def run_processing():
    excel_path = excel_path_var.get()
    output_dir = output_dir_var.get()

    if not os.path.isfile(excel_path):
        messagebox.showerror("Lỗi", "Vui lòng chọn file Excel hợp lệ.")
        return
    if not os.path.isdir(output_dir):
        messagebox.showerror("Lỗi", "Vui lòng chọn thư mục lưu kết quả hợp lệ.")
        return

    status_label.config(text="Đang xử lý, vui lòng chờ...")
    root.update_idletasks()

    try:
        results = process_thalassemia_excel(excel_path, output_dir)
        messagebox.showinfo("Hoàn tất", f"Đã tạo {len(results)} file báo cáo.")
    except Exception as e:
        messagebox.showerror("Lỗi khi xử lý", str(e))
    finally:
        status_label.config(text="Xử lý xong!")

    config = {
    "last_excel_path": excel_path_var.get(),
    "last_output_dir": output_dir_var.get()
    }
    save_config(config)

root = tk.Tk()
root.title("Tạo báo cáo Thalassemia")

excel_path_var = tk.StringVar()
output_dir_var = tk.StringVar()

config = load_config()
excel_path_var.set(config.get("last_excel_path", ""))
output_dir_var.set(config.get("last_output_dir", ""))

tk.Label(root, text="1. Chọn file Excel:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
tk.Entry(root, textvariable=excel_path_var, width=50).grid(row=0, column=1, padx=10)
tk.Button(root, text="Tìm", command=select_excel).grid(row=0, column=2, padx=5)

tk.Label(root, text="2. Chọn thư mục lưu kết quả:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
tk.Entry(root, textvariable=output_dir_var, width=50).grid(row=1, column=1, padx=10)
tk.Button(root, text="Tìm", command=select_output_folder).grid(row=1, column=2, padx=5)

tk.Button(root, text="Tạo báo cáo", command=run_processing, bg="green", fg="white").grid(row=2, column=1, pady=20)

status_label = tk.Label(root, text="", fg="blue")
status_label.grid(row=3, column=1, pady=5)

root.mainloop()