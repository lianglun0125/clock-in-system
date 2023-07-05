import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import csv
import datetime
import os
import requests
import json
import getpass
import sys

# Version 1.3.0

# 檢查資料夾是否存在，如果不存在，則創建資料夾
folder_path = 'C:/Program Files (x86)/Work-Record'
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

def download_file(url, save_path):
    response = requests.get(url)
    with open(save_path, 'wb') as file:
        file.write(response.content)

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
        #print(f"Latest version: {latest_version}")
        #print(f"Current version: {current_version}")

        # 進行版本比較
        result = compare_versions(current_version, latest_version)
        if result == -1:
            response = messagebox.askyesno('有新版本', '有新版本可用，是否要更新？')      
            if response:
                username = getpass.getuser()
                #print(username)

                # 下載最新版本的 Clock-In-NEW.exe 檔案
                download_path = f'C:/Users/{username}/Desktop'
                save_path = os.path.join(download_path, file_name)
                download_file(download_url, save_path)
                messagebox.showinfo('成功', '下載完成，即將關閉舊版本。')
                root.destroy()

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
        record = [date, '', time]  

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
        date = record[0]
        signin = record[1]
        signout = record[2]

        if date not in records:
            records[date] = {'日期': date, '簽到時間': '', '簽退時間': ''}
        
        if signin:
            records[date]['簽到時間'] = signin
        
        if signout:
            records[date]['簽退時間'] = signout

    export_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV 檔案', '*.csv')])

    if export_path:
        with open(export_path, 'w', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=['日期', '簽到時間', '簽退時間'])
            writer.writeheader()
            writer.writerows(records.values())

        messagebox.showinfo('成功', f'簽到記錄已匯出至 {export_path}。')

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

root = tk.Tk()
root.title('1.3.0')
root.geometry("205x165")
root.resizable(False, False)

signin_button = ttk.Button(root, text='今日簽到', command=lambda: save_record('簽到'))
signin_button.pack()

signout_button = ttk.Button(root, text='今日簽退', command=lambda: save_record('簽退'))
signout_button.pack()

history_button = ttk.Button(root, text='查看簽到歷史紀錄', command=show_history)
history_button.pack()

export_records_button = ttk.Button(root, text='匯出簽到紀錄', command=export_records)
export_records_button.pack()

check_update_button = ttk.Button(root, text='檢查更新', command=check_update)
check_update_button.pack()

piglabel = tk.Label(root, text='阿咪今天簽到了嗎？', font=('微軟正黑體', 10, 'bold'), fg='#191970')
piglabel.pack()

updatelabel = tk.Label(root, text='Last Update - 20230706', font=('微軟正黑體', 10, 'bold'), fg='#191970')
updatelabel.pack()

exe_name = os.path.basename(sys.argv[0])
if exe_name == "Clock-In-NEW.exe":
        new_name = "Clock-In.exe"
        if os.path.exists("Clock-In.exe"):
            os.remove("Clock-In.exe")
        os.rename(exe_name, new_name)
        messagebox.showinfo('提示', '首次執行此版本，環境設置成功，請重新啟動。')
        sys.exit()
else:
        root.mainloop()