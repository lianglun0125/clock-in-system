import ctypes
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter.ttk import Progressbar
import csv
import datetime
import os
import requests
import json
import getpass
import sys
from threading import Thread


# Version 1.3.5

# 檢查資料夾是否存在，如果不存在，則創建資料夾
folder_path = 'C:/Program Files (x86)/Work-Record'
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

def download_file(url, save_path, progress_var, percentage_var):
    response = requests.get(url, stream=True)
    total_length = int(response.headers.get('content-length'))
    dl = 0
    with open(save_path, 'wb') as file:
        for chunk in response.iter_content(chunk_size=1024):
            if chunk:
                dl += len(chunk)
                file.write(chunk)
                progress = int((dl / total_length) * 100)
                progress_var.set(progress)
                percentage_var.set(f'{progress}%')

def compare_versions(current_version, latest_version):
    current_parts = current_version.split('.')
    latest_parts = latest_version.split('.')

    # 將版本號的每個部分轉換為整數，並補齊位數不足的部分
    current_parts = [int(part) for part in current_parts]
    latest_parts = [int(part) for part in latest_parts]

    # 使用 zip 函式將兩個版本號的對應部分進行比較
    for current_part, latest_part in zip(current_parts, latest_parts):
        if current_part < latest_part:
            return -1  # 最新版本大於當前版本
        elif current_part > latest_part:
            return 2  # 最新版本小於當前版本

    # 如果迴圈結束後還沒有返回，則表示版本號的前部分相同，比較版本號的長度
    if len(current_parts) < len(latest_parts):
        return -1  # 最新版本大於當前版本
    elif len(current_parts) > len(latest_parts):
        return 2  # 最新版本小於當前版本

    return 0  # 兩個版本號相同

def check_update():
    # 指定你的 GitHub 使用者名稱、專案名稱和檔案名稱
    username = 'lianglun0125'
    repo = 'clock-in-system'
    file_name = 'Clock-In-NEW.exe'

    # 請求 GitHub API 獲取最新的發佈版本資訊
    releases_url = f'https://api.github.com/repos/{username}/{repo}/releases/latest'
    response = requests.get(releases_url)
    data = json.loads(response.text)
    #print(data)

    # 獲取最新版本的下載 URL
    download_url = None
    for asset in data['assets']:
        if asset['name'] == file_name:
            download_url = asset['browser_download_url']
            break

    if download_url:
        latest_version = data['tag_name']
        current_version = root.title()

        result = compare_versions(current_version, latest_version)
        if result == -1:
            response = messagebox.askyesno('有新版本', '有新版本可用，是否要更新？')
            if response:
                username = getpass.getuser()
                download_path = f'C:/Users/{username}/Desktop'
                save_path = os.path.join(download_path, file_name)

                download_window = tk.Toplevel(root)
                download_window.title('新版本下載進度')
                download_window.resizable(False,False)

                progress_var = tk.IntVar()
                percentage_var = tk.StringVar()

                progress_bar = Progressbar(download_window, length=300, mode='determinate', variable=progress_var)
                progress_bar.pack(pady=2)

                percentage_label = tk.Label(download_window, textvariable=percentage_var)
                percentage_label.pack()

                download_thread = Thread(target=download_file, args=(download_url, save_path, progress_var, percentage_var))
                download_thread.start()

                def check_download():
                    if download_thread.is_alive():
                        root.after(100, check_download)
                    else:
                        download_window.destroy()
                        messagebox.showinfo('成功', '下載完成，即將關閉舊版本。')
                        root.destroy()

                root.after(100, check_download)

        elif result == 0:
            messagebox.showinfo('訊息', '已經是最新版本。')
        elif result == 2:
            messagebox.showinfo('錯誤', '版本錯誤，請聯繫作者。')
        else:
            messagebox.showinfo('訊息', '已經是最新版本。')
    else:
        pass

def save_record(action):
    date = datetime.date.today()
    time = datetime.datetime.now().strftime("%H:%M:%S")

    if action == '簽到':
        record = [date, time, ''] 
    elif action == '簽退':
        # 檢查是否已簽到，如果沒有則顯示警告訊息
        with open('C:/Program Files (x86)/Work-Record/attendance.csv', 'r') as file:
            reader = csv.reader(file)
            data = list(reader)
        
        for record in data:
            if record[0] == str(date) and record[1]:  # 如果今天已經簽到
                record = [date, '', time]
                break
        else:
            messagebox.showwarning('錯誤', '阿咪今天尚未簽到，無法簽退哦。')
            return

    with open('C:/Program Files (x86)/Work-Record/attendance.csv', 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(record)

    messagebox.showinfo('成功', f'今日{action}時間已成功紀錄。')

def export_records():
    with open('C:/Program Files (x86)/Work-Record/attendance.csv', 'r') as file:
        reader = csv.reader(file)
        data = list(reader)

    records = {}
    for record in data:
        if len(record) >= 3:
            date = record[0]
            signin = record[1]
            signout = record[2]

            if date not in records:
                records[date] = {'日期': date, '簽到時間': '', '簽退時間': '', '工時': ''}
            
            if signin:
                if not records[date]['簽到時間'] or signin < records[date]['簽到時間']:
                    records[date]['簽到時間'] = signin
            
            if signout:
                if not records[date]['簽退時間'] or signout > records[date]['簽退時間']:
                    records[date]['簽退時間'] = signout

    # 計算工時
    for record in records.values():
        signin_time = record['簽到時間']
        signout_time = record['簽退時間']
        
        if signin_time and signout_time:
            work_hours = calculate_work_hours(signin_time, signout_time)
            record['工時'] = work_hours

    export_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV 檔案', '*.csv')])

    if export_path:
        with open(export_path, 'w', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=['日期', '簽到時間', '簽退時間', '工時'])
            writer.writeheader()
            writer.writerows(records.values())

        messagebox.showinfo('成功', f'簽到記錄已匯出至 {export_path}。')

def calculate_work_hours(signin_time, signout_time):

    signin = datetime.datetime.strptime(signin_time, '%H:%M:%S')
    signout = datetime.datetime.strptime(signout_time, '%H:%M:%S')
    work_duration = signout - signin
    work_hours = work_duration.total_seconds() / 3600  # 轉換為小時數
    work_hours_str = f'{work_hours:.2f} 小時'  # 格式化工時字串，取小數點後兩位
    return work_hours_str

def show_history():
    def filter_records():
        selected_year = year_var.get()
        selected_month = month_var.get()

        if selected_year == '' or selected_month == '':
            messagebox.showerror('錯誤', '請記得選年、月再按搜尋')
            return

        with open('C:/Program Files (x86)/Work-Record/attendance.csv', 'r') as file:
            reader = csv.reader(file)
            data = list(reader)

        filtered_data = [record for record in data if record[0].startswith(f'{selected_year}-{selected_month}')]

        table.delete(*table.get_children())

        prev_date = ''
        for record in filtered_data:
            if record[0] == prev_date:
                table.item(prev_item, values=(prev_values[0], prev_values[1], record[2]))
            else:
                prev_item = table.insert('', 'end', values=record)
                prev_values = record
                prev_date = record[0]

    history_window = tk.Toplevel(root)
    history_window.title('歷史簽到記錄')
    history_window.geometry("265x335")
    history_window.resizable(False, False)

    year_var = tk.StringVar()
    month_var = tk.StringVar()

    year_label = ttk.Label(history_window, text='年')
    year_label.pack()
    year_combobox = ttk.Combobox(history_window, textvariable=year_var, values=[''] + list(range(2023, 2028)), state='readonly')
    year_combobox.pack()

    month_label = ttk.Label(history_window, text='月')
    month_label.pack()
    month_combobox = ttk.Combobox(history_window, textvariable=month_var, values=[''] + ["{:02d}".format(i) for i in range(1, 13)], state='readonly')
    month_combobox.pack()

    filter_button = ttk.Button(history_window, text='搜尋', command=filter_records)
    filter_button.pack()

    columns = ('日期', '簽到時間', '簽退時間')
    table = ttk.Treeview(history_window, columns=columns, show='headings')

    table.column('日期', minwidth=80, width=80)
    table.column('簽到時間', minwidth=80, width=80)
    table.column('簽退時間', minwidth=80, width=80)

    for col in columns:
        table.heading(col, text=col)

    scrollbar = ttk.Scrollbar(history_window, orient=tk.VERTICAL, command=table.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    table.configure(yscrollcommand=scrollbar.set)
    table.pack()

    def center_align(event):
        table.column("#0", anchor="center")
        for col in columns:
            table.column(col, anchor="center")

    table.bind("<Configure>", center_align)

    root.update()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if True:
    root = tk.Tk()
    root.title('1.3.5')
    root.geometry("205x180")
    root.resizable(False, False)

    signin_button = ttk.Button(root, text='今日簽到', command=lambda: save_record('簽到'))
    signin_button.pack(pady=3)

    signout_button = ttk.Button(root, text='今日簽退', command=lambda: save_record('簽退'))
    signout_button.pack(pady=3)

    history_button = ttk.Button(root, text='查看簽到紀錄', command=show_history)
    history_button.pack(pady=3)

    export_history_button = ttk.Button(root, text='匯出簽到紀錄', command=export_records)
    export_history_button.pack(pady=3)

    check_update_button = ttk.Button(root, text='檢查更新', command=check_update)
    check_update_button.pack(pady=3)

    piglabel = tk.Label(root, text='阿咪今天簽到了嗎？', font=('微軟正黑體', 12, 'bold'), fg='#191970')
    piglabel.pack(pady=3)

    exe_name = os.path.basename(sys.argv[0])
    if exe_name == "Clock-In-NEW.exe":
        new_name = "Clock-In.exe"
        if os.path.exists("Clock-In.exe"):
            os.remove("Clock-In.exe")
        os.rename(exe_name, new_name)
        messagebox.showinfo('提示', '首次執行此版本，環境設置成功，請使用管理員身分重新啟動。')
        sys.exit()
    else:
        root.mainloop()
else:
    messagebox.showerror('Permission Denied', '阿咪，要用管理員身分執行哦')
    sys.exit()
