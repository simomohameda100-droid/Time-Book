import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import openpyxl
import os
from datetime import datetime
from plyer import notification

# =========================
# 1️⃣ إعداد ملف Excel
# =========================
file_name = "missions.xlsx"
if not os.path.exists(file_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Missions"
    ws.append(["Mission", "Date", "Time"])
    ws_finished = wb.create_sheet("Finished")
    ws_finished.append(["Mission", "Date", "Time", "Finished At"])
    wb.save(file_name)

# =========================
# 2️⃣ حفظ مهمة جديدة
# =========================
def save_mission():
    mission = entry_mission.get()
    date = entry_date.get()
    time = entry_time.get()

    if not mission or not date or not time:
        messagebox.showwarning("⚠️ Warning", "Please fill all fields!")
        return

    wb = openpyxl.load_workbook(file_name)
    ws = wb["Missions"]
    ws.append([mission, date, time])
    wb.save(file_name)

    entry_mission.delete(0, tk.END)
    messagebox.showinfo("✅ Saved", f"Mission '{mission}' added successfully!")
    load_missions()

# =========================
# 3️⃣ تحميل المهام
# =========================
def load_missions():
    for item in tree.get_children():
        tree.delete(item)

    wb = openpyxl.load_workbook(file_name)
    ws = wb["Missions"]

    for row in ws.iter_rows(min_row=2, values_only=True):
        tree.insert("", tk.END, values=row)

# =========================
# 4️⃣ حذف مهمة
# =========================
def delete_mission():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("⚠️ Warning", "Select a mission to delete!")
        return

    mission = tree.item(selected_item, "values")[0]

    wb = openpyxl.load_workbook(file_name)
    ws = wb["Missions"]

    for row in ws.iter_rows(min_row=2):
        if row[0].value == mission:
            ws.delete_rows(row[0].row)
            break

    wb.save(file_name)
    tree.delete(selected_item)

# =========================
# 5️⃣ تعديل مهمة
# =========================
def edit_mission():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("⚠️ Warning", "Select a mission to edit!")
        return

    values = tree.item(selected_item, "values")
    entry_mission.delete(0, tk.END)
    entry_mission.insert(0, values[0])
    entry_date.set_date(values[1])
    entry_time.delete(0, tk.END)
    entry_time.insert(0, values[2])

    btn_save_edit.config(state="normal")
    btn_add.config(state="disabled")

def save_edit():
    selected_item = tree.selection()
    if not selected_item:
        return

    new_mission = entry_mission.get()
    new_date = entry_date.get()
    new_time = entry_time.get()

    wb = openpyxl.load_workbook(file_name)
    ws = wb["Missions"]

    for row in ws.iter_rows(min_row=2):
        if row[0].value == tree.item(selected_item, "values")[0]:
            row[0].value = new_mission
            row[1].value = new_date
            row[2].value = new_time
            break

    wb.save(file_name)
    load_missions()

    btn_save_edit.config(state="disabled")
    btn_add.config(state="normal")

# =========================
# 6️⃣ إشعار عند وقت المهمة
# =========================
def check_notifications():
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    wb = openpyxl.load_workbook(file_name)
    ws = wb["Missions"]
    ws_finished = wb["Finished"]

    for row in list(ws.iter_rows(min_row=2, values_only=False)):
        mission, date, time = row[0].value, row[1].value, row[2].value
        if f"{date} {time}" == now:
            notification.notify(
                title="⏰ Mission Reminder",
                message=f"Mission: {mission}\nDate: {date} {time}",
                timeout=10
            )
            ws_finished.append([mission, date, time, datetime.now().strftime("%Y-%m-%d %H:%M")])
            ws.delete_rows(row[0].row)

    wb.save(file_name)
    load_missions()
    root.after(60000, check_notifications)  # كل دقيقة

# =========================
# 7️⃣ عرض المهام المنجزة
# =========================
def show_finished_missions():
    wb = openpyxl.load_workbook(file_name)
    if "Finished" not in wb.sheetnames:
        messagebox.showinfo("ℹ️ Info", "No finished missions yet!")
        return

    ws_finished = wb["Finished"]

    finished_win = tk.Toplevel(root)
    finished_win.title("📚 Finished Missions")
    finished_win.geometry("500x400")
    finished_win.configure(bg="#f0f8ff")

    tk.Label(finished_win, text="📚 Finished Missions", font=("Arial", 14, "bold"),
             bg="#4682B4", fg="white", pady=5).pack(fill="x")

    text_area = tk.Text(finished_win, font=("Arial", 12))
    text_area.pack(fill="both", expand=True, padx=10, pady=10)

    for row in ws_finished.iter_rows(min_row=2, values_only=True):
        mission, date, time, finished_at = row
        text_area.insert(tk.END, f"✅ {mission} | 📅 {date} | 🕒 {time} | Finished: {finished_at}\n")

    text_area.config(state="disabled")

# =========================
# 8️⃣ واجهة التطبيق
# =========================
root = tk.Tk()
root.title("📌 Mission Reminder App")
root.geometry("700x600")
root.configure(bg="#f0f8ff")

# إدخال المهام
frame_input = tk.Frame(root, bg="#f0f8ff")
frame_input.pack(pady=10)

tk.Label(frame_input, text="Mission:", bg="#f0f8ff", font=("Arial", 12)).grid(row=0, column=0, padx=5, pady=5)
entry_mission = tk.Entry(frame_input, font=("Arial", 12), width=30)
entry_mission.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_input, text="Date:", bg="#f0f8ff", font=("Arial", 12)).grid(row=1, column=0, padx=5, pady=5)
entry_date = DateEntry(frame_input, width=27, background='darkblue', foreground='white', borderwidth=2, date_pattern="yyyy-mm-dd")
entry_date.grid(row=1, column=1, padx=5, pady=5)

tk.Label(frame_input, text="Time (HH:MM):", bg="#f0f8ff", font=("Arial", 12)).grid(row=2, column=0, padx=5, pady=5)
entry_time = tk.Entry(frame_input, font=("Arial", 12), width=30)
entry_time.grid(row=2, column=1, padx=5, pady=5)

# أزرار التحكم
frame_buttons = tk.Frame(root, bg="#f0f8ff")
frame_buttons.pack(pady=10)

btn_add = ttk.Button(frame_buttons, text="➕ Add Mission", command=save_mission)
btn_add.grid(row=0, column=0, padx=10)

ttk.Button(frame_buttons, text="🗑 Delete", command=delete_mission).grid(row=0, column=1, padx=10)
ttk.Button(frame_buttons, text="✏ Edit", command=edit_mission).grid(row=0, column=2, padx=10)
btn_save_edit = ttk.Button(frame_buttons, text="💾 Save Edit", command=save_edit, state="disabled")
btn_save_edit.grid(row=0, column=3, padx=10)
ttk.Button(frame_buttons, text="📚 Show Finished", command=show_finished_missions).grid(row=0, column=4, padx=10)

# جدول عرض المهام
frame_table = tk.Frame(root, bg="#f0f8ff")
frame_table.pack(pady=20, fill="both", expand=True)

columns = ("Mission", "Date", "Time")
tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=10)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150, anchor="center")
tree.pack(fill="both", expand=True)

# تحميل المهام عند البدء
load_missions()
check_notifications()

root.mainloop()
