import json
import re
import time
import tkinter.ttk as ttk
import webbrowser
from pprint import pprint
from time import sleep
from tkinter import *

import pymongo
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# 分別依照系級分類，取得系級以下的科/部字典清單
def get_division_dict():
    division_dict = {}

    for td in soup_clinic_table.find_all("td", valign="top"):
        for img in td.find_all("img"):
            if img["alt"] != "*":
                division_name = img["alt"]
                division_dict[division_name] = []
                table = img.find_next_sibling("table")
                for a in table.find_all("a"):
                    if a.text:
                        division_dict[division_name].append(a.text)
    return division_dict


def callback_select_division(event):
    value = str(combo_select_division.get())
    combo_select_department.config(values=["請選擇"] + division_dict[value])
    combo_select_department.current(0)


# input:選定的科別 output:選定科別底下的門診清單
def find_clinic(url):
    browser.get(url)  # 前往指定科別網址
    soup = BeautifulSoup(browser.page_source, "html.parser")

    clinic_list = []
    rules = soup.find_all(rules="all")  # 上午/下午/晚上
    for a in range(len(rules)):
        trs = rules[a].select("tr")
        trs.pop(0)
        for tr in trs:
            clinic = tr.select("span")[0].text
            if clinic not in clinic_list and "1" not in clinic:
                clinic_list.append(clinic)
    return clinic_list


def callback_select_department(event):
    for td in soup_clinic_table.find_all("td", align="left"):
        if combo_select_department.get() == td.a.string:
            url = "https://reg.ntuh.gov.tw/webadministration/" + td.a["href"]
            break
    combo_select_clinic.config(values=["請選擇"] + find_clinic(url))
    combo_select_clinic.current(0)


# treeview的雙擊事件
def take_me_to_register_main(event):
    select_id = tree_main.selection()
    values = tree_main.item(select_id, "values")
    link = values[3]
    if "RegistForm.aspx" in link:
        # for opening the link in browser
        webbrowser.open(link)


# treeview的雙擊事件(for搜尋醫生名)
def take_me_to_register_doctor(event):
    select_id = tree_doctor.selection()
    values = tree_doctor.item(select_id, "values")
    link = values[4]
    if "RegistForm.aspx" in link:
        webbrowser.open(link)


# 選定門診後，尋找本周與下周可預約的醫師
def find_available_appointment_by_clinic():
    items = tree_main.get_children()
    [tree_main.delete(item) for item in items]

    url_list = [browser.current_url]
    if browser.find_element(By.ID, "LinkNextWeekContent"):
        next_week_js = browser.find_element(By.ID, "LinkNextWeekContent").get_attribute(
            "href"
        )  # 取得下一周的連結內的 javascript
        browser.execute_script(next_week_js)
        url_list.append(browser.current_url)  # 下周的 URL
        browser.back()

    all_available = []
    for url in url_list:
        browser.get(url)
        soup = BeautifulSoup(browser.page_source, "html.parser")
        for table in soup.find_all(rules="all"):  # 0是上午 1是下午
            trs = table.select("tr")
            for tr in trs:
                if combo_select_clinic.get() == tr.select("span")[0].text:
                    doctors = tr.select("a")
                    for doctor in doctors:
                        if doctor.has_attr("onmouseover"):
                            browser.execute_script(doctor["href"])
                            soup = BeautifulSoup(browser.page_source, "html.parser")
                            update_available_appointment(
                                soup,
                                all_available,
                                "DoctorServiceListInSeveralDaysTemplateIDSE_GridViewDoctorServiceList",
                            )
                            browser.back()

    all_available_sort_by_time = sorted(
        all_available, key=lambda d: strptime(d["time"])
    )  # 用 time 欄位排序整個 all_available[]
    # pprint(all_available_sort_by_time)
    for data in all_available_sort_by_time:
        # print(data)
        tree_main.insert(
            "",
            "end",
            text=data["appointment_type"],
            values=[data["name"], data["time"], data["note"], data["link"]],
        )
    if not all_available_sort_by_time:
        tree_main.insert("", 0, text="查無可掛號門診")


# 輸入醫生名稱搜尋門診資訊
def find_available_appointment_by_doctor_name():
    doctor_name = entry_doctor_name.get()  # 取得輸入的醫生名稱

    browser.get(
        "https://reg.ntuh.gov.tw/WebAdministration/DoctorServiceQueryByDrName.aspx?HospCode=T0&QueryName="
        + doctor_name
    )
    soup = BeautifulSoup(browser.page_source, "html.parser")

    items = tree_doctor.get_children()
    [tree_doctor.delete(item) for item in items]

    all_available = []

    update_available_appointment(
        soup,
        all_available,
        "DoctorServiceListInSeveralDaysInput_GridViewDoctorServiceList",
    )

    all_available_sort_by_time = sorted(
        all_available, key=lambda d: strptime(d["time"])
    )  # 用 time 欄位排序整個 all_available[]
    # pprint(all_available_sort_by_time)
    for data in all_available_sort_by_time:
        # print(data)
        tree_doctor.insert(
            "",
            "end",
            text=data["appointment_type"],
            values=[
                data["name"],
                data["time"],
                data["clinic"],
                data["note"],
                data["link"],
            ],
        )

    if not all_available_sort_by_time:
        tree_doctor.insert("", 0, text="查無可掛號門診")


# 將可預約門診時間轉換成可比較的格式
def strptime(data_time):
    # 確認是否符合 "111.5.16(五) 上午" 格式
    if not re.match(
        r"\d{1,3}\.\d{1,2}\.\d{1,2}\([\u4e00-\u9fa5]\) +[\u4e00-\u9fa5]{2}", data_time
    ):
        return ""
    year, month, day = re.findall(r"\d{1,3}", data_time)
    am_pm = 0 if re.findall(r"[\u4e00-\u9fa5]{2}", data_time)[0] == "上午" else 1
    return f"{year}-{month.zfill(2)}-{day.zfill(2)}-{am_pm}"


# 取得可預約門診資訊
def update_available_appointment(soup, all_available, table_id):
    doctor_services = soup.find_all("table", id=table_id)[0].select("tr")
    for doctor_service in doctor_services:
        if len(doctor_service.select("td")):
            if not doctor_service.select("td")[0].find("a"):  # 第一格沒有 <a> 代表無法預約
                continue
            js_href = doctor_service.select("td")[0].find("a")[
                "href"
            ]  # 取得掛號的 javascript
            browser.execute_script(js_href)
            time.sleep(0.01)
            appointment_link = browser.current_url
            browser.back()
            doctor_available = {
                "name": doctor_service.select("td")[1].find("span").text,  # 醫事人員
                "time": doctor_service.select("td")[2].find("span").text,  # 門診時間
                "hospital": doctor_service.select("td")[3].find("span").text,  # 院區
                "location": doctor_service.select("td")[4].find("span").text,  # 地點
                "clinic": doctor_service.select("td")[5].find("span").text,  # 診別
                "clinic_number": doctor_service.select("td")[7].find("span").text,  # 診別
                "appointment_type": doctor_service.select("td")[0]
                .find("a")
                .text,  # 掛號欄位的 text, e.g., 初診/掛號
                "link": appointment_link,  # 掛號的超連結
                "note": doctor_service.select("td")[9].text.strip("\n"),  # 備註
            }

            if doctor_available not in all_available:
                #                 print(doctor_available)
                all_available.append(doctor_available)


def search_doctor_name(*args):
    value = var.get()
    if not value:
        combobox_doctor_name_search["values"] = []
        return
    doctor_name_list = get_keyword_result(str(value))
    # 醫院網頁搜尋陳開頭的醫師，會出現內分泌新陳代謝科此項目
    if doctor_name_list[0] == "內分泌新陳代謝科":
        doctor_name_list.remove("內分泌新陳代謝科")
    # print(doctor_name_list)
    if len(doctor_name_list) == 0:
        combobox_doctor_name_search["values"] = []
        return
    combobox_doctor_name_search["values"] = doctor_name_list


def get_session_jsessionid():
    # 取得 cookies
    res = session.get(url="https://www.ntuh.gov.tw/ntuh/FindDr.action")
    cookies = res.cookies.get_dict()
    if not cookies:
        print(f"Cookies 取得錯誤。 status_code:{res.status_code}")
        return []
    return cookies["JSESSIONID"]


def get_keyword_result(keyword):
    headers = {"cookie": f"JSESSIONID={jsessionid}"}
    url = f"https://www.ntuh.gov.tw/ntuh/FindDrAjax!autocomplete.action?query={keyword}"
    res = session.get(url, headers=headers)
    try:
        result = res.json()
    except json.JSONDecodeError as e:
        print(f"JSONDecodeError: json 格式解析錯誤。 status_code: {res.status_code}")
        return []
    time.sleep(1)
    return result["suggest"]


def search_comment_from_db():
    # 清空treeview所有的row
    items = tree_comment.get_children()
    [tree_comment.delete(item) for item in items]

    doctor_name = combobox_doctor_name_search.get()
    search_result = get_keyword_result(doctor_name)

    if len(search_result) == 0 or (
        len(search_result) == 1 and search_result[0] != doctor_name
    ):  # 若台大系統沒有該醫生時
        tree_comment.insert(
            "", 0, text="", value=[f'"{doctor_name}" 醫師似乎不存在於該醫院中。', "", ""]
        )
        return
    if len(search_result) > 1:  # 若台大系統查詢該關鍵字有多位醫生時
        tree_comment.insert(
            "", 0, text="", value=[f'關鍵字 "{doctor_name}" 有多位醫師，請再試著詳細得輸入。', "", ""]
        )
        return

    result = list(collection.find({"name": doctor_name}))  # 用醫生姓名去 db 查詢評論
    if len(result) == 0:  # 若我們的 DB 沒有該醫生的評論時
        tree_comment.insert(
            "", 0, text="", value=[f"{doctor_name} 醫師尚未有任何評論，請在下方留言後送出。", "", ""]
        )
        return

    for row in result:
        tree_comment.insert(
            "", "end", text=row["name"], values=[row["comment"], row["time"]]
        )
        print(row)
        # 這邊要一筆一筆插入醫生評論


def send_comment_to_db():
    doctor_name = combobox_doctor_name_search.get()
    search_result = get_keyword_result(doctor_name)
    if len(entry_comment.get()) < 1:
        return
    if len(search_result) == 1:
        # 插入評論到 DB
        new_comment = {
            "name": doctor_name,
            "comment": entry_comment.get(),
            "time": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
        }
        collection.insert_one(new_comment)
    entry_comment.delete(0, "end")
    search_comment_from_db()


# 初始化
def init():
    global browser, collection, session, jsessionid, soup_clinic_table, division_dict

    # 設定 webdriver
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # 啟動 Headless 無頭模式
    chrome_options.add_argument("--disable-gpu")  # 關閉 GPU 避免某些系統或是網頁出錯
    browser = webdriver.Chrome(
        "web_driver\\chromedriver.exe", chrome_options=chrome_options
    )

    # 設定 mongodb 連線
    mongo_client = pymongo.MongoClient("mongodb://0.tcp.ap.ngrok.io:11761/")
    db = mongo_client["bigdataDB"]
    collection = db["doctor_comment"]

    # 設定 requests session
    session = requests.session()
    jsessionid = get_session_jsessionid()

    # 總院: https://reg.ntuh.gov.tw/WebAdministration/ClinicTable.aspx?var=T0
    # 系清單>科清單>門診(普通門診....)
    res = requests.get(
        "https://reg.ntuh.gov.tw/WebAdministration/ClinicTable.aspx?var=T0"
    )
    soup_clinic_table = BeautifulSoup(res.text)

    # 取得門診清單
    division_dict = get_division_dict()


init()

root = Tk()
root.title("醫院掛號查詢")
root.geometry("690x400")

style = ttk.Style(root)
style.theme_create(
    "custom_tab",
    parent="alt",
    settings={
        "TNotebook.Tab": {
            "configure": {"padding": [7, 2]},
        }
    },
)
style.theme_use("clam")  #'default','classic','clam','custom_tab'
my_notebook = ttk.Notebook()  # 分頁欄
my_notebook.pack()
# 製作框架
frame_main = Frame(my_notebook, width=600, height=400)
frame_doctor = Frame(my_notebook, width=600, height=400)
frame_comment = Frame(my_notebook, width=600, height=400)

frame_main.pack(fill="both", expand=1)
frame_doctor.pack(fill="both", expand=1)
frame_comment.pack(fill="both", expand=1)

my_notebook.add(frame_main, text="依科別查詢")
my_notebook.add(frame_doctor, text="依醫師名查詢")
my_notebook.add(frame_comment, text="醫師評價")

# ---------------------frame_main----------------------------------------
label_select_division = Label(frame_main, text="選擇科系")
label_select_division.grid(column=0, row=0)
combo_select_division = ttk.Combobox(
    frame_main, values=["請選擇"] + list(division_dict.keys()), state="readonly"
)
combo_select_division.grid(column=0, row=1)
combo_select_division.current(0)
combo_select_division.bind("<<ComboboxSelected>>", callback_select_division)

label_select_department = Label(frame_main, text="選擇部")
label_select_department.grid(column=1, row=0)
combo_select_department = ttk.Combobox(frame_main, values=["請選擇"], state="readonly")
combo_select_department.grid(column=1, row=1)
combo_select_department.current(0)
combo_select_department.bind("<<ComboboxSelected>>", callback_select_department)

label_select_clinic = Label(frame_main, text="選擇門診")
label_select_clinic.grid(column=2, row=0, pady=5)
combo_select_clinic = ttk.Combobox(frame_main, values=["請選擇"], state="readonly")
combo_select_clinic.grid(column=2, row=1)
combo_select_clinic.current(0)

btn_search_main = Button(frame_main)
btn_search_main.config(
    width=4, height=2, text="查詢", command=find_available_appointment_by_clinic
)
btn_search_main.grid(column=3, row=0, rowspan=2)

tree_main = ttk.Treeview(frame_main)
tree_main.bind("<Double-Button-1>", take_me_to_register_main)
tree_main["columns"] = ("1", "2", "3")
tree_main.column("#0", width=120, minwidth=25)
tree_main.column("1", anchor="w", width=120)
tree_main.column("2", anchor="w", width=200)
tree_main.column("3", anchor="w", width=240)
# tree.column("4",anchor='w',width=160)
tree_main.heading("#0", text="掛號", anchor="w")
tree_main.heading("1", text="醫事人員", anchor="w")
tree_main.heading("2", text="門診時間", anchor="w")
tree_main.heading("3", text="備註", anchor="w")
# tree.heading("4",text="連結",anchor='w')
tree_main.grid(column=0, row=2, columnspan=4)

label_only_available = Label(frame_main, text="* 僅顯示可預約門診")
label_only_available.config(fg="#FF0000")
label_only_available.grid(column=0, row=3)

# --------------------------------------------------------------------


# ---------------------frame_doctor----------------------------------------
entry_doctor_name = Entry(frame_doctor)
entry_doctor_name.insert(0, "醫生姓名")  # 預設提示字
entry_doctor_name.grid(column=2, row=0, ipadx=0, ipady=5)

btn_search_doctor = Button(frame_doctor)
btn_search_doctor.config(
    width=4, height=2, text="查詢", command=find_available_appointment_by_doctor_name
)
btn_search_doctor.grid(column=3, row=0)

tree_doctor = ttk.Treeview(frame_doctor)
tree_doctor.bind("<Double-Button-1>", take_me_to_register_doctor)
tree_doctor["columns"] = ("1", "2", "3", "4")
tree_doctor.column("#0", width=120, minwidth=25)
tree_doctor.column("1", anchor="w", width=70)
tree_doctor.column("2", anchor="w", width=120)
tree_doctor.column("3", anchor="w", width=90)
tree_doctor.column("4", anchor="w", width=240)
tree_doctor.heading("#0", text="掛號", anchor="w")
tree_doctor.heading("1", text="醫事人員", anchor="w")
tree_doctor.heading("2", text="門診時間", anchor="w")
tree_doctor.heading("3", text="科別", anchor="w")
tree_doctor.heading("4", text="備註", anchor="w")
tree_doctor.grid(column=0, row=2, columnspan=4)
# --------------------------------------------------------------------

# ---------------------frame_comment----------------------------------------
var = StringVar(value="")
var.trace_add("write", search_doctor_name)
label_doctor_name = Label(frame_comment, text="醫師名稱")
label_doctor_name.grid(column=0, row=0, columnspan=4)
combobox_doctor_name_search = ttk.Combobox(frame_comment, textvariable=var)
combobox_doctor_name_search.grid(column=1, row=0)

btn_search_comment = Button(frame_comment)
btn_search_comment.config(
    width=4, height=2, text="查詢\n評價", command=search_comment_from_db
)
btn_search_comment.grid(column=2, row=0)

tree_comment = ttk.Treeview(frame_comment)
tree_comment["columns"] = ("1", "2")
tree_comment.column("#0", width=110, minwidth=25)
tree_comment.column("1", anchor="w", width=420)
tree_comment.column("2", anchor="w", width=130)
tree_comment.heading("#0", text="醫師姓名", anchor="w")
tree_comment.heading("1", text="評論", anchor="w")
tree_comment.heading("2", text="時間", anchor="w")
tree_comment.grid(column=0, row=2, columnspan=4)

entry_comment = Entry(frame_comment)
entry_comment.grid(column=0, row=3, ipady=20, ipadx=120)

btn_send_comment = Button(frame_comment)
btn_send_comment.config(width=4, height=2, text="送出", command=send_comment_to_db)
btn_send_comment.grid(column=3, row=3, columnspan=2, rowspan=2)
# --------------------------------------------------------------------

root.mainloop()

