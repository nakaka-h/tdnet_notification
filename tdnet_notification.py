import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

import sys
from bs4 import BeautifulSoup
import requests
import pandas as pd
from pyquery import PyQuery as pq
import json
import schedule
from time import sleep
import datetime
import threading
import queue
import re
from plyer import notification
import webbrowser

import urllib3
from urllib3.exceptions import InsecureRequestWarning
urllib3.disable_warnings(InsecureRequestWarning)

# GUI
root =tk.Tk()
root.geometry("500x400")
root.resizable(False, False)
root.title("適時開示Watcher")
#root.iconbitmap(default='./W.ico')

# 日本時間
t_delta = datetime.timedelta(hours=9)
JST = datetime.timezone(t_delta, 'JST')
now = datetime.datetime.now(JST)
today = now.strftime('%Y%m%d')

# TdnetのURL
Tdnet_url = "https://www.release.tdnet.info/inbs/I_main_00.html"

# 開示件数取得用URL
Tdnet_kaijiSum_url = "https://www.release.tdnet.info/inbs/I_list_001_" + today + ".html"

# 登録銘柄保存用jsonファイル
settings_json = "./stocks.json"

# 登録銘柄保存用リスト
with open(settings_json, "r") as g:
    settings = json.load(g)
    kabucodes = settings["kabucodes"]

# 適時開示通知リスト
notices = []

# 適時開示URLリスト
URLs = []

# 通知初期設定
notice_bool = tk.BooleanVar()

# 現在時刻
#now = datetime.datetime.now()
#now_time = now.strftime('%H:%M:%S')
#print(now_time)

# 更新の有無に応じた関数を実行
def task():
    # 現在の最終更新日時
    with open('update_time.txt') as f:
        old_elem = f.read()

    # 新たな最終更新日時の取得
    res_update = requests.get(Tdnet_url)
    soup_update = BeautifulSoup(res_update.content, 'html.parser')
    new_elem = str(soup_update.select('#last-update > div'))
    res_kaijiSum = requests.get(Tdnet_kaijiSum_url)
    soup_kaijiSum = BeautifulSoup(res_kaijiSum.content, 'html.parser')
    kaijiSum = str(soup_kaijiSum.select('#pager-box-top > div.kaijiSum'))

    # 更新の有無を表示
    if new_elem == old_elem:
        pass

    else:
        with open('update_time.txt', 'w') as f:
            f.write(new_elem)
            last_update = new_elem.replace("[<div>最終更新日時：", "").replace("年", "-").replace("月", "-").replace("日", "").replace("</div>]", ":00")
            sleep(1)

            # 開示件数を取得してスクレイピングするURLを作成
            if kaijiSum == "[]":
                pass
            else:
                kaijiSum = kaijiSum.split(" ")
                kaijiSum = kaijiSum[2]
                kaijiSum = int(kaijiSum.replace('全', '').replace('件</div>]', ''))

                pages = -(-kaijiSum // 100)

                now = datetime.datetime.now(JST)
                today = now.strftime('%Y%m%d')
                last_update_hm = now.strftime('%H:%M')

                kabucodes = [s + '0' for s in kabucodes]

                linklist = []

                for i in range(1, pages+1):
                    if i < 10:
                        page_number = "00" + str(i)
                    else:
                        page_number = "0" + str(i)

                    Tdnet_mainlist_url = "https://www.release.tdnet.info/inbs/I_list_" + page_number + "_" + today + ".html"

                    if i == 1:
                        df = pd.read_html(Tdnet_mainlist_url, encoding='utf-8')[3]
                    else:
                        df = pd.concat([df, pd.read_html(Tdnet_mainlist_url, encoding='utf-8')[3]])

                    page = pq(Tdnet_mainlist_url, encoding='utf-8')

                    for a in page('a'):
                        link = pq(a).attr('href')
                        if ".zip" in link:
                            pass
                        else:
                            link = 'https://www.release.tdnet.info/inbs/' + link
                            linklist.append(link)

                # 証券コードごとに更新があるか確認、あれば通知
                df['開示URL'] = linklist
                df = df.rename(columns={df.columns[0]: '時刻', df.columns[1]: 'コード', df.columns[2]: '会社名', df.columns[3]: '表題'})

                df_last_update = df[df['時刻'].str.contains(last_update_hm)]

                df_noticelist = df_last_update[df_last_update['コード'].astype(str).str.contains("|".join(kabucodes))]

                for i in range(0, len(df_noticelist['時刻'])):
                    df_noticelist = df_noticelist.rename(index={df_noticelist.index[i]: i})

                df_empty = df_noticelist.empty

                if df_empty == True:
                    pass

                else:
                    for i in range(0, len(df_noticelist['時刻'])):
                        notice_sentence = last_update+" / "+df_noticelist.loc[i, 'コード'].astype(str)+" / "+df_noticelist.loc[i, '会社名']+" / "+df_noticelist.loc[i, '表題']
                        lb2.insert(0, notice_sentence)
                        notices.insert(0, notice_sentence)
                        URLs.insert(0, df_noticelist.loc[i, '開示URL'])
                        if notice_bool.get() == True:
                            notification.notify(title=df_noticelist.loc[i, 'コード'].astype(str)+"  "+df_noticelist.loc[i, '会社名'], message="適時開示があります", app_name="適時開示Watcher", app_icon='./W.ico', timeout = 60)



##メニューバーの関数

# 環境設定ボタン

# 終了ボタン
def quit():
    root.destroy()

# 銘柄追加ボタン
def open_add_window():
    dialog = tk.Toplevel()
    dialog.title("銘柄追加    - IRWatcher")
    dialog.geometry("200x300")
    dialog.resizable(False, False)
    dialog.grab_set()

    # 登録銘柄一覧表示
    add_window_frame1 = tk.Frame(dialog, padx = 10)
    add_window_frame1.place(x = 10, y = 80)
    kabulist = tk.StringVar(value=kabucodes)
    lb = tk.Listbox(add_window_frame1, listvariable=kabulist, width = 25, height = 11, selectmode=tk.MULTIPLE)
    lb.grid(row = 0, column = 0, sticky=(tk.E, tk.W))

    scrollbar = tk.Scrollbar(add_window_frame1, orient=tk.VERTICAL, command=lb.yview)
    lb['yscrollcommand'] = scrollbar.set
    scrollbar.grid(row = 0, column = 1, sticky=(tk.N, tk.S))

    # 銘柄追加画面ウィジェットの作成
    add_window_label1 = tk.Label(dialog, text="銘柄追加")
    add_window_input1 = ttk.Entry(dialog, width = 10)
    # 追加ボタン
    def add():
        if re.match(re.compile('[0-9]+'), add_window_input1.get()):
            if len(add_window_input1.get()) == 4:
                lb.insert(tk.END, add_window_input1.get())
                kabucodes.append(add_window_input1.get())
                add_window_input1.delete(0, tk.END)
                kabucodes_forjson = {"kabucodes": kabucodes}
                with open(settings_json, "w") as g:
                    json.dump(kabucodes_forjson, g)
            else:
                messagebox.showerror("エラー", "4桁の証券コードを入力してください")
        else:
            messagebox.showerror("エラー", "半角数字を入力してください")

    # 削除ボタン
    def delete():
        delete_bool = messagebox.askyesno("確認", "選択中の証券コードを本当に削除しますか？")
        if delete_bool == True:
            deleted_index = list(lb.curselection())
            for i in sorted(deleted_index, reverse=True):
                lb.delete(i)
            for i in sorted(deleted_index, reverse=True):
                kabucodes.pop(i)
            kabucodes_forjson = {"kabucodes": kabucodes}
            with open(settings_json, "w") as g:
                json.dump(kabucodes_forjson, g)

    add_window_button1 = ttk.Button(dialog, width = 6, text="追加", command=add)
    add_window_label2 = tk.Label(dialog, text="登録銘柄一覧")
    add_window_button2 = ttk.Button(dialog, width = 6, text="削除", command=delete)#警告ポップアップ
    add_window_button3 = ttk.Button(dialog, width = 6, text="閉じる", command=dialog.destroy)

    # 銘柄追加画面ウィジェットの設置
    add_window_label1.place(x = 10, y = 20)
    add_window_input1.place(x = 70, y = 20)
    add_window_button1.place(x = 150, y = 18)
    add_window_label2.place(x = 10, y = 50)
    add_window_button2.place(x = 150, y = 50)
    add_window_button3.place(x = 150, y = 270)

# メニューバーの作成・設置
menubar = tk.Menu(root)
root.config(menu=menubar)

setting_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="設定", menu=setting_menu)

setting_menu.add_command(label="銘柄追加", command=open_add_window)
setting_menu.add_checkbutton(label="通知", variable = notice_bool, onvalue=True, offvalue=False)
setting_menu.add_command(label="終了", command=quit)

# 状態の起動時デフォルト表示
state = tk.StringVar()
state.set("停止中")

# キュー作成
q_running_bool = queue.Queue()

# 更新監視関数定期実行
def running(running_bool, q_running_bool):
    running_bool = q_running_bool.get()
    # 0~5秒からschduletask開始
    while True:
        now = datetime.datetime.now()
        if now.strftime("%S") >= "00" and now.strftime("%S") <= "03":
            stopbutton.config(state='enable')
            # task関数を1秒毎に定期実行
            schedule.every(1).seconds.do(task)
            break

    while running_bool:
        schedule.run_pending()
        #sleep(1)
        running_bool = q_running_bool.get()
        q_running_bool.put(True)

# 実行ボタン関数
def run():
    state.set("実行中")
    label1.config(fg = '#3eb370')
    startbutton.config(state='disable')
    schedule.clear()
    running_bool = True
    q_running_bool.put(running_bool)
    q_running_bool.put(running_bool)
    thread = threading.Thread(target=running, args=(running_bool, q_running_bool))
    thread.start()

# 停止ボタン関数
def stop():
    state.set("停止中")
    label1.config(fg = '#e60033')
    stopbutton.config(state='disable')
    startbutton.config(state='enable')
    running_bool = False
    q_running_bool.put(running_bool)

# ハイパーリンク関数
def jumpURL(event):
    if noticelist == []:
        pass
    else:
        webbrowser.open_new(URLs[lb2.curselection()[0]])



# 適時開示通知一覧表示
frame1 = tk.Frame(root, padx = 10)
frame1.place(x = 10, y = 250)
noticelist = tk.StringVar(value=notices)
lb2 = tk.Listbox(frame1, listvariable=noticelist, width = 75, height = 7)
lb2.grid(row = 0, column = 0, sticky=(tk.E, tk.W))

scrollbar2 = tk.Scrollbar(frame1, orient=tk.VERTICAL, command=lb2.yview)
lb2['yscrollcommand'] = scrollbar2.set
scrollbar2.grid(row = 0, column = 1, sticky=(tk.N, tk.S))

scrollbar3 = tk.Scrollbar(frame1, orient=tk.HORIZONTAL, command=lb2.xview)
lb2['xscrollcommand'] = scrollbar3.set
scrollbar3.grid(row = 1, column = 0, sticky=(tk.W, tk.E))

lb2.bind('<<ListboxSelect>>', jumpURL)

# メイン画面ウィジェットの作成
label1 = tk.Label(root, font=("MSゴシック", "50", "bold"), fg = '#e60033', textvariable=state)
startbutton =ttk.Button(root, width = 10, padding = 5, text="実行", command=run)
stopbutton = ttk.Button(root, width = 10, padding = 5, text="停止", state='disabled', command=stop)
label2 = tk.Label(root, text="適時開示一覧")

# メイン画面ウィジェットの設置
label1.place(x = 150, y = 50)
startbutton.place(x = 175, y = 175)
stopbutton.place(x = 260, y = 175)
label2.place(x = 20, y = 225)

root.mainloop()
