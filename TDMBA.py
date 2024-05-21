import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pandas as pd
import shutil
import os
import win32print
import win32ui
import win32con
from tkinter import filedialog
from persiantools.jdatetime import JalaliDateTime as jdatetime


people = []
arzabkoonandeh = []
next_chest_strap = 1  # شروع از 1
evaluators = []
interview_types = {
    "مصاحبه 1": ["مصاحبه 1"],
    "مصاحبه 2": ["مصاحبه 2"]
}
def validate_number(input):
    if input.isdigit() or input == "":
        return True
    else:
        return False

def validate_meli_length(input):
    if len(input) > 10:
        return False
    return True

def clear_entries():
    entry_first_name.delete(0, 'end')
    entry_last_name.delete(0, 'end')
    entry_meli.delete(0, 'end')
    entry_phone.delete(0, 'end')

def add_person():
    global next_chest_strap
    first_name = entry_first_name.get()
    last_name = entry_last_name.get()
    meli = entry_meli.get()
    phone = entry_phone.get()

    if not first_name or not last_name or not meli or not phone:
        messagebox.showwarning("خطا", "لطفاً مشخصات ارزیاب شونده را به صورت کامل وارد کنید.")
        return

    if len(meli) != 10:
        messagebox.showwarning("خطا", "کد ملی باید 10 رقم باشد.")
        return

    if len(phone) != 11:
        messagebox.showwarning("خطا", "شماره همراه باید 11 رقم باشد.")
        return

    for person in people:
        if person['کد ملی'] == meli:
            messagebox.showwarning("خطا", "کد ملی تکراری است.")
            return

    person = {
        'شماره': next_chest_strap,
        'نام': first_name,
        'نام_خانوادگی': last_name,
        'کد ملی': meli,
        'شماره همراه': phone,
        'روند روز': {
            'مصاحبه 1': False,
            'مصاحبه 2': False,
            ' نوشتاری': False,
            ' گروهی': False
        }
    }
    people.append(person)
    next_chest_strap += 1
    list_people()
    clear_entries()

def list_people():
    tree.delete(*tree.get_children())
    for i, person in enumerate(people, start=1):
        interviewers = []
        for task, done in person['روند روز'].items():
            if done and task in ["مصاحبه 1", "مصاحبه 2"]:
                for evaluator in evaluators:
                    if evaluator["دوره مصاحبه"] == task:
                        interviewers.append(f"{task}: {evaluator['نام']}")
        tree.insert("", "end", text=str(i), values=(
            person['شماره'], person['نام'], person['نام_خانوادگی'], person['کد ملی'], person['شماره همراه'],
            ", ".join(interviewers) if interviewers else "ندارد"
        ))

def edit_person_tasks():
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("خطا", "لطفاً یک ارزیاب شونده را انتخاب کنید.")
        return

    person_index = int(tree.item(selected_item)['text']) - 1
    person = people[person_index]

    top = tk.Toplevel()
    top.title(f"ویرایش وظایف {person['نام']} {person['نام_خانوادگی']}")

    task_vars = {}
    for task, done in person['روند روز'].items():
        var = tk.BooleanVar(value=done)
        task_vars[task] = var
        chk = ttk.Checkbutton(top, text=task, variable=var, command=lambda t=task: update_label(t, task_vars))
        chk.pack(pady=5)

        if task in ['مصاحبه 1', 'مصاحبه 2']:
            btn_interview = tk.Button(top, text=f"انتخاب نوع {task}", command=lambda t=task: select_evaluator(t, task_vars))
            btn_interview.pack(pady=5)

    selected_tasks_label = tk.Label(top, text="روند طی شده : ")
    selected_tasks_label.pack(pady=10)

    def update_label(task, task_vars):
        selected_tasks = [t for t, var in task_vars.items() if var.get()]
        selected_tasks_label.config(text="روند طی شده : " + ", ".join(selected_tasks))

    def show_interview_options(person, task):
        option_top = tk.Toplevel()
        option_top.title(f"انتخاب نوع {task}")

        def select_option(option):
            person['روند روز'][task] = option
            option_top.destroy()
            list_people()

        btn_inperson = tk.Button(option_top, text="مصاحبه حضوری", command=lambda: select_option("مصاحبه حضوری"))
        btn_inperson.pack(pady=5)

        btn_phone = tk.Button(option_top, text="مصاحبه تلفنی", command=lambda: select_option("مصاحبه تلفنی"))
        btn_phone.pack(pady=5)

    def save_tasks():
        for task, var in task_vars.items():
            person['روند روز'][task] = var.get()
        list_people()
        top.destroy()

    save_btn = tk.Button(top, text="ذخیره روند طی شده", command=save_tasks)
    save_btn.pack(pady=10)

def get_selected_candidates(evaluator_name):
    selected_candidates = []
    for person in people:
        if person['روند روز'].get('مصاحبه 1') == evaluator_name or person['روند روز'].get('مصاحبه 2') == evaluator_name:
            selected_candidates.append(f"{person['نام']} {person['نام_خانوادگی']}")
    return selected_candidates

def list_evaluators():
    top = tk.Toplevel()
    top.title("لیست ارزیاب کننده‌ها")

    evaluator_tree = ttk.Treeview(top, columns=("نام", "دوره مصاحبه", "ارزیاب شونده‌ها"), show="headings")
    evaluator_tree.heading("نام", text="نام")
    evaluator_tree.heading("دوره مصاحبه", text="دوره مصاحبه")
    evaluator_tree.heading("ارزیاب شونده‌ها", text="ارزیاب شونده‌ها")

    for i, evaluator in enumerate(evaluators, start=1):
        selected_candidates = get_selected_candidates(evaluator["نام"])
        evaluator_tree.insert("", "end", text=str(i), values=(evaluator["نام"], evaluator["دوره مصاحبه"], ", ".join(selected_candidates)))

    evaluator_tree.pack(pady=20, padx=20)

def add_evaluator():
    top = tk.Toplevel()
    top.title("اضافه کردن ارزیاب کننده")

    tk.Label(top, text="نام:").grid(row=0, column=0)
    name_entry = tk.Entry(top)
    name_entry.grid(row=0, column=1)

    tk.Label(top, text="نوع مصاحبه:").grid(row=1, column=0)
    interview_round_var = tk.StringVar()
    interview_round_combobox = ttk.Combobox(top, textvariable=interview_round_var, values=["مصاحبه 1", "مصاحبه 2"])
    interview_round_combobox.grid(row=1, column=1)

    def save_evaluator():
        evaluator = {
            "نام": name_entry.get(),
            "دوره مصاحبه": interview_round_var.get()
        }
        evaluators.append(evaluator)
        top.destroy()

    save_btn = tk.Button(top, text="ذخیره", command=save_evaluator)
    save_btn.grid(row=2, columnspan=2)

def show_completed_tasks():
    top = tk.Toplevel()
    top.title("ارزیابی‌های امروز")

    completed_tree = ttk.Treeview(top, columns=("شماره", "روند طی شده امروز", "نام", "نام خانوادگی", "ارزیاب و مصاحبه"), show="headings")
    completed_tree.heading("شماره", text="شماره")
    completed_tree.heading("روند طی شده امروز", text="روند طی شده امروز")
    completed_tree.heading("نام", text="نام")
    completed_tree.heading("نام خانوادگی", text="نام خانوادگی")
    completed_tree.heading("ارزیاب و مصاحبه", text="ارزیاب و مصاحبه")

    for person in people:
        completed_tasks = [task for task, done in person['روند روز'].items() if done]
        completed_tasks_str = ", ".join(completed_tasks)

        interviewers = []
        for task in ["مصاحبه 1", "مصاحبه 2"]:
            if task in completed_tasks:
                for evaluator in evaluators:
                    if evaluator["دوره مصاحبه"] == task:
                        interviewers.append(f"{task}: {evaluator['نام']}")
        interviewer_str = ", ".join(interviewers) if interviewers else "ندارد"

        completed_tree.insert("", "end", values=(person['شماره'], completed_tasks_str, person['نام'], person['نام_خانوادگی'], interviewer_str))

    completed_tree.pack(pady=20, padx=20)


    def print_completed_tasks():
        tasks_to_print = []
        for person in people:
            completed_tasks = [task for task, done in person['روند روز'].items() if done]
            if completed_tasks:
                tasks_to_print.append(f"{person['نام']} {person['نام_خانوادگی']}: {', '.join(completed_tasks)}")

        text_to_print = "\n".join(tasks_to_print)
        printer_name = win32print.GetDefaultPrinter()
        hprinter = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(hprinter, 2)
        printer_info['pDevMode'].Orientation = win32con.DMORIENT_LANDSCAPE if len(tasks_to_print) > 10 else win32con.DMORIENT_PORTRAIT
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc("Tasks")
        hdc.StartPage()
        hdc.DrawText(text_to_print, (0, 0, 1000, 1000), win32con.DT_LEFT)
        hdc.EndPage()
        hdc.EndDoc()

    btn_print = tk.Button(top, text="چاپ اطلاعات", command=print_completed_tasks)
    btn_print.pack(pady=10)


def select_evaluator(task, task_vars):
    def on_select():
        selected_evaluator = evaluator_listbox.get(tk.ACTIVE)
        for key, value in interview_types.items():
            if selected_evaluator in value:
                task_vars[task].set(key)
                break
        top.destroy()

    top = tk.Toplevel()
    top.title(f"انتخاب ارزیاب برای {task}")

    evaluator_listbox = tk.Listbox(top, selectmode=tk.SINGLE)
    evaluator_listbox.pack(pady=10)

    for evaluator in evaluators:
        evaluator_listbox.insert(tk.END, evaluator["نام"])

    btn_select = tk.Button(top, text="انتخاب", command=on_select)
    btn_select.pack(pady=10)


root = tk.Tk()
root.title("روند ارزیابی")

frame_entry = tk.Frame(root)
frame_entry.pack(pady=20)

tk.Label(frame_entry, text="نام:").grid(row=1, column=0)
entry_first_name = tk.Entry(frame_entry)
entry_first_name.grid(row=1, column=1)

tk.Label(frame_entry, text="نام خانوادگی:").grid(row=2, column=0)
entry_last_name = tk.Entry(frame_entry)
entry_last_name.grid(row=2, column=1)

tk.Label(frame_entry, text="کد ملی:").grid(row=3, column=0)
entry_meli = tk.Entry(frame_entry, validate="key", validatecommand=[(root.register(validate_number), "%P"), (root.register(validate_meli_length), "%P")])
entry_meli.grid(row=3, column=1)

tk.Label(frame_entry, text="شماره همراه:").grid(row=4, column=0)
entry_phone = tk.Entry(frame_entry, validate="key", validatecommand=(root.register(validate_number), "%P"))
entry_phone.grid(row=4, column=1)

btn_add = tk.Button(frame_entry, text="اضافه کردن ارزیاب شونده", command=add_person)
btn_add.grid(row=5, column=0, columnspan=2, pady=10)

frame_tree = tk.Frame(root)
frame_tree.pack(pady=20)

tree = ttk.Treeview(frame_tree, columns=("شماره", "نام", "نام خانوادگی", "کد ملی", "شماره همراه"), show="headings")
tree.heading("شماره", text="شماره")
tree.heading("نام", text="نام")
tree.heading("نام خانوادگی", text="نام خانوادگی")
tree.heading("کد ملی", text="کد ملی")
tree.heading("شماره همراه", text="شماره همراه")
tree.pack()

btn_edit_tasks = tk.Button(root, text="ویرایش روند ارزیاب شونده", command=edit_person_tasks)
btn_edit_tasks.pack(pady=10)

btn_show_completed = tk.Button(root, text="روند ارزیابی ها امروز", command=show_completed_tasks)
btn_show_completed.pack(pady=10)

btn_add_evaluator = tk.Button(root, text="اضافه کردن ارزیاب کننده", command=add_evaluator)
btn_add_evaluator.pack(pady=10)

btn_list_evaluators = tk.Button(root, text="نمایش ارزیاب‌کننده‌ها", command=list_evaluators)
btn_list_evaluators.pack(pady=10)


#*******************

# تعریف متغیرهای گلوبال


# تابع پشتیبان‌گیری
def backup_to_csv(people):
    # پوشه‌ی مشخص شده برای ذخیره بک‌آپها
    backup_dir = "D:\\Arzyabi"

    # ایجاد یک DataFrame از لیست افراد
    df = pd.DataFrame(people)

    # تاریخ و ساعت شمسی را بدست می‌آوریم
    now = jdatetime.now()  # از JalaliDateTime استفاده می‌کنیم
    shamsi_date_time = now.strftime('%Y%m%d_%H%M%S')

    # نام فایل
    file_name = f"backup_{shamsi_date_time}.csv"

    # مسیر ذخیره‌سازی
    file_path = os.path.join(backup_dir, file_name)

    try:
        # ذخیره‌ی DataFrame به عنوان یک فایل CSV
        df.to_csv(file_path, index=False, encoding='utf-8-sig')
        messagebox.showinfo("پشتیبان‌گیری", f"پشتیبان‌گیری با موفقیت انجام شد. {file_name}")
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در پشتیبان‌گیری: {e}")

# در قسمت اضافه کردن دکمه‌ها:
btn_backup = tk.Button(root, text="پشتیبان‌گیری", command=backup_to_csv)
btn_backup.pack(pady=10)


# تابع بازیابی
def restore_from_csv():
    global people, arzabkoonandeh

    file_path = filedialog.askopenfilename(title="بازیابی بک‌آپ", filetypes=[("CSV Files", "*.csv")])
    if file_path:
        try:
            df = pd.read_csv(file_path, encoding='utf-8')

            if 'ارزیاب کننده ها' in df.columns:
                df['روند روز'] = df['روند روز'].apply(eval)

                arzabkoonandeh = df['ارزیاب کننده ها'].tolist()
                df = df.drop(columns=['ارزیاب کننده ها'])

                people = df.to_dict('records')
                messagebox.showinfo("اطلاعات", "بازیابی با موفقیت انجام شد.")
                list_people()
            else:
                messagebox.showerror("خطا", "فایل بازیابی شده دارای ستون 'ارزیاب کننده ها' نیست.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در بازیابی بک‌آپ: {e}")


# دکمه بازیابی
btn_restore_backup = tk.Button(root, text="بازیابی بک‌آپ", command=restore_from_csv)
btn_restore_backup.pack(pady=10)


def on_closing():
    if messagebox.askokcancel("خروج", "آیا از بستن برنامه مطمئن هستید؟"):
        backup_to_csv(people)
        jnow = jdatetime.now()  # اینجا تعریف کنید
        shamsi_date = jnow.strftime("%Y%m%d")
        backup_dir = f"D:\\Arzyabi\\backup_{shamsi_date}"
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        shutil.copy(__file__, os.path.join(backup_dir, f"backup_{shamsi_date}.py"))
        root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
