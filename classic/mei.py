# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_name(value):
    return clean_text(value).replace("（", "(").replace("）", ")").replace(" ", "")


def find_column(columns, names):
    normalized = {str(c).strip().lower(): c for c in columns}
    for name in names:
        if name.lower() in normalized:
            return normalized[name.lower()]
    return None


def get_supplier_name(filename):
    stem = Path(filename).stem.strip()

    # 与原版一致：存在下划线时，取第一个下划线前的内容。
    if "_" in stem:
        return stem.split("_", 1)[0].strip()

    # PDF 询证函一般直接以完整公司名称命名。
    for suffix in ("往来及交易询证函", "询证函", "对账单"):
        if stem.endswith(suffix):
            stem = stem[:-len(suffix)].rstrip(" _-—")
            break

    return stem


def create_and_send_emails():
    folder_path = folder_entry.get().strip()
    contact_file = contact_entry.get().strip()

    if not folder_path or not contact_file:
        messagebox.showwarning("提示", "请先选择文件夹和通讯录！")
        return

    try:
        df_contacts = pd.read_excel(contact_file, dtype=object)
        df_contacts.columns = [str(c).strip() for c in df_contacts.columns]

        name_col = find_column(df_contacts.columns, ["名称", "供应商名称", "公司名称"])
        to_col = find_column(df_contacts.columns, ["邮箱", "收件人", "收件邮箱", "电子邮箱"])
        cc_col = find_column(df_contacts.columns, ["抄送", "抄送人", "抄送邮箱"])

        if name_col is None or to_col is None:
            messagebox.showerror(
                "错误",
                "通讯录表格中必须包含“名称”和“邮箱”两列！\n"
                "也兼容“名称 / 收件人 / 抄送人”格式。",
            )
            return

        email_dict = {}
        for _, row in df_contacts.iterrows():
            name = clean_text(row[name_col])
            if not name:
                continue

            to_email = clean_text(row[to_col])
            cc_email = clean_text(row[cc_col]) if cc_col is not None else ""

            email_dict[normalize_name(name)] = {
                "name": name,
                "to": to_email,
                "cc": cc_email,
            }

        outlook = win32.Dispatch("outlook.application")
        count = 0
        skip_count = 0

        for filename in os.listdir(folder_path):
            extension = os.path.splitext(filename)[1].lower()
            if extension not in (".xlsx", ".xls", ".xlsm", ".pdf"):
                continue

            supplier_name = get_supplier_name(filename)
            contact_info = email_dict.get(normalize_name(supplier_name))

            # 若文件名还带有其他文字，尝试以通讯录公司名包含匹配。
            if not contact_info:
                normalized_stem = normalize_name(Path(filename).stem)
                matches = [
                    (len(company_name), info)
                    for company_name, info in email_dict.items()
                    if company_name and company_name in normalized_stem
                ]
                matches.sort(key=lambda item: item[0], reverse=True)
                if matches:
                    contact_info = matches[0][1]
                    supplier_name = contact_info["name"]

            if not contact_info or not contact_info["to"]:
                print(f"找不到【{supplier_name}】的收件邮箱，已跳过。")
                skip_count += 1
                continue

            mail = outlook.CreateItem(0)

            # 与原版一致：先打开邮件窗口，让 Outlook 自动加载默认签名。
            mail.Display()
            mail.To = contact_info["to"]

            if contact_info["cc"]:
                mail.CC = contact_info["cc"]

            if extension == ".pdf":
                mail.Subject = f"【请查收】{supplier_name} - 往来及交易询证函"
                my_text = """
                <div style="font-family: 'Microsoft YaHei', SimSun, sans-serif; font-size: 14px;">
                你好！<br><br>
                附件为往来及交易询证函，请查收并协助核对。<br>
                如信息无误，请签章确认；如有不符，请列明不符项目及具体内容。<br>
                完成后烦请将签章后的文件回传，谢谢！<br><br>
                </div>
                """
            else:
                current_month = datetime.now().month
                mail.Subject = f"【请查收】{supplier_name} - {current_month}月份对账单"
                my_text = """
                <div style="font-family: 'Microsoft YaHei', SimSun, sans-serif; font-size: 14px;">
                你好！<br><br>
                附件本月对账单请查收，请尽快核对，如无误请开票。<br><br>
                新电子发票要求：<br><br>
                1）仅需提供OFD和XML档，PDF档不需要<br>
                2）发票命名：网上直接下载后的文件名称不用改动，在下载后的文件名前加贵公司简称提供。<br>
                3）截止收票日期：每月15号<br><br>
                </div>
                """

            mail.HTMLBody = my_text + mail.HTMLBody

            attachment_path = os.path.abspath(os.path.join(folder_path, filename))
            mail.Attachments.Add(attachment_path)
            count += 1

        messagebox.showinfo(
            "处理完成",
            f"操作结束！\n成功生成了 {count} 封邮件。\n"
            f"跳过了 {skip_count} 个没有收件邮箱的供应商。",
        )

    except Exception as exc:
        messagebox.showerror("运行出错", f"发生错误:\n{exc}")


def select_folder():
    path = filedialog.askdirectory(title="选择存放对账单或询证函的文件夹")
    if path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, path)


def select_contact_file():
    filepath = filedialog.askopenfilename(
        title="选择供应商通讯录",
        filetypes=[("Excel 文件", "*.xlsx;*.xls;*.xlsm")],
    )
    if filepath:
        contact_entry.delete(0, tk.END)
        contact_entry.insert(0, filepath)


root = tk.Tk()
root.title("Outlook 自动对账单发送工具 (保留签名版)")
root.geometry("550x200")


tk.Label(root, text="附件文件夹: ").grid(
    row=0, column=0, padx=10, pady=20, sticky="e"
)
folder_entry = tk.Entry(root, width=45)
folder_entry.grid(row=0, column=1, padx=10, pady=20)
tk.Button(root, text="选择文件夹...", command=select_folder).grid(
    row=0, column=2, padx=10, pady=20
)


tk.Label(root, text="供应商通讯录: ").grid(
    row=1, column=0, padx=10, pady=10, sticky="e"
)
contact_entry = tk.Entry(root, width=45)
contact_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="选择表格...", command=select_contact_file).grid(
    row=1, column=2, padx=10, pady=10
)


tk.Button(
    root,
    text="▶ 生成邮件预览",
    bg="#0078D4",
    fg="white",
    font=("Arial", 12, "bold"),
    command=create_and_send_emails,
).grid(row=2, column=1, pady=20)


root.mainloop()
