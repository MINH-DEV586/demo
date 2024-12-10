import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
from datetime import datetime
import pandas as pd

# Hàm lưu dữ liệu vào file CSV
def save_data():
    data = {
        "Mã NV": entry_ma.get(),
        "Tên": entry_ten.get(),
        "Ngày sinh": entry_ngaysinh.get(),
        "Giới tính": gender_var.get(),
        "Số CMND": entry_cmnd.get(),
        "Ngày cấp": entry_ngaycap.get(),
        "Nơi cấp": entry_noicap.get(),
        "Chức danh": entry_chucdanh.get()
    }

    if "" in data.values():
        messagebox.showwarning("Cảnh báo", "Vui lòng nhập đầy đủ thông tin!")
        return

    # Ghi vào file CSV
    try:
        with open("nhanvien.csv", "a", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(file, fieldnames=data.keys())
            if file.tell() == 0:  # Nếu file rỗng, ghi header
                writer.writeheader()
            writer.writerow(data)
        messagebox.showinfo("Thành công", "Dữ liệu đã được lưu!")
        clear_fields()
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể lưu dữ liệu: {e}")

# Hàm xóa các trường nhập liệu
def clear_fields():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    entry_ngaysinh.delete(0, tk.END)
    entry_cmnd.delete(0, tk.END)
    entry_ngaycap.delete(0, tk.END)
    entry_noicap.delete(0, tk.END)
    entry_chucdanh.delete(0, tk.END)
    gender_var.set("Nam")

# Hàm hiển thị danh sách nhân viên có sinh nhật hôm nay
def show_birthdays_today():
    try:
        today = datetime.today().strftime("%d/%m")
        employees = []
        with open("nhanvien.csv", "r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row["Ngày sinh"][:5] == today:
                    employees.append(row)

        if employees:
            result = "\n".join([f"{emp['Tên']} - {emp['Mã NV']}" for emp in employees])
            messagebox.showinfo("Sinh nhật hôm nay", result)
        else:
            messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay!")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "File dữ liệu không tồn tại!")

# Hàm xuất danh sách ra file Excel
def export_to_excel():
    try:
        df = pd.read_csv("nhanvien.csv", encoding="utf-8")
        df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y")
        df["Tuổi"] = (datetime.now() - df["Ngày sinh"]).dt.days // 365
        df = df.sort_values(by="Tuổi", ascending=False)

        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            df.to_excel(filepath, index=False, encoding="utf-8")
            messagebox.showinfo("Thành công", f"Danh sách đã được xuất ra {filepath}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất dữ liệu: {e}")

# Tạo giao diện Tkinter
root = tk.Tk()
root.title("Quản lý nhân viên")

# Các trường nhập liệu
frame = ttk.Frame(root, padding=10)
frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

ttk.Label(frame, text="Mã NV:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
entry_ma = ttk.Entry(frame)
entry_ma.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame, text="Tên:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
entry_ten = ttk.Entry(frame)
entry_ten.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame, text="Ngày sinh (DD/MM/YYYY):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
entry_ngaysinh = ttk.Entry(frame)
entry_ngaysinh.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(frame, text="Giới tính:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
gender_var = tk.StringVar(value="Nam")
gender_frame = ttk.Frame(frame)
gender_frame.grid(row=3, column=1, padx=5, pady=5)
ttk.Radiobutton(gender_frame, text="Nam", variable=gender_var, value="Nam").grid(row=0, column=0)
ttk.Radiobutton(gender_frame, text="Nữ", variable=gender_var, value="Nữ").grid(row=0, column=1)

ttk.Label(frame, text="Số CMND:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
entry_cmnd = ttk.Entry(frame)
entry_cmnd.grid(row=4, column=1, padx=5, pady=5)

ttk.Label(frame, text="Ngày cấp (DD/MM/YYYY):").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
entry_ngaycap = ttk.Entry(frame)
entry_ngaycap.grid(row=5, column=1, padx=5, pady=5)

ttk.Label(frame, text="Nơi cấp:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
entry_noicap = ttk.Entry(frame)
entry_noicap.grid(row=6, column=1, padx=5, pady=5)

ttk.Label(frame, text="Chức danh:").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
entry_chucdanh = ttk.Entry(frame)
entry_chucdanh.grid(row=7, column=1, padx=5, pady=5)

# Các nút chức năng
btn_frame = ttk.Frame(root, padding=10)
btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))

ttk.Button(btn_frame, text="Lưu", command=save_data).grid(row=0, column=0, padx=5, pady=5)
ttk.Button(btn_frame, text="Sinh nhật hôm nay", command=show_birthdays_today).grid(row=0, column=1, padx=5, pady=5)
ttk.Button(btn_frame, text="Xuất danh sách", command=export_to_excel).grid(row=0, column=2, padx=5, pady=5)

# Chạy ứng dụng
root.mainloop()
