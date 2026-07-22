# -*- coding: utf-8 -*-
import os, re, traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import win32com.client as win32

EXTS={'.pdf','.xls','.xlsx','.xlsm'}
def s(v): return '' if pd.isna(v) else str(v).strip()
def norm(v): return re.sub(r'[\s\u3000]+','',s(v).replace('（','(').replace('）',')')).casefold().rstrip('._-—–')
def col(cols,names):
 m={re.sub(r'[\s\u3000]+','',str(x)).casefold():x for x in cols}
 for n in names:
  k=re.sub(r'[\s\u3000]+','',n).casefold()
  if k in m:return m[k]

def read_contacts(path):
 df=pd.read_excel(path,dtype=object); df.columns=[s(x) for x in df.columns]
 nc=col(df.columns,['名称','供应商名称','公司名称']); tc=col(df.columns,['收件人','邮箱','收件邮箱','电子邮箱','email','e-mail']); cc=col(df.columns,['抄送人','抄送','抄送邮箱','cc'])
 if nc is None or tc is None: raise ValueError('通讯录必须包含“名称”和“收件人”列；也兼容“名称/邮箱/抄送”。')
 out={}; ci=df.columns.get_loc(cc) if cc is not None else None
 for _,r in df.iterrows():
  name=s(r[nc]); to=s(r[tc])
  if not name:continue
  ccs=[]
  if ci is not None:
   for i in range(ci,len(df.columns)):
    v=s(r.iloc[i])
    if v and '@' in v:ccs.append(v)
  out[norm(name)]={'name':name,'to':to,'cc':';'.join(dict.fromkeys(ccs))}
 return out

def candidates(filename):
 stem=Path(filename).stem.strip(); vals=[stem]
 for sep in ('_','-','—','–'):
  if sep in stem:vals.append(stem.split(sep,1)[0].strip())
 for p in (r'[\s_-]*(?:\d{4}年)?\d{1,2}月(?:份)?对账单$',r'[\s_-]*(?:往来及交易)?询证函$',r'[\s_-]*对账单$'):
  x=re.sub(p,'',stem,flags=re.I).strip()
  if x:vals.append(x)
 return vals

def match(filename,contacts):
 for x in candidates(filename):
  if norm(x) in contacts:return contacts[norm(x)]
 st=norm(Path(filename).stem); hits=[(len(k),v) for k,v in contacts.items() if k and k in st]; hits.sort(reverse=True,key=lambda x:x[0])
 return hits[0][1] if hits and (len(hits)==1 or hits[0][0]>hits[1][0]) else None

def mail_text(c,path,month):
 if path.suffix.lower()=='.pdf':
  return f'【请查收】{c["name"]} - 往来及交易询证函',"""<div style="font-family:'Microsoft YaHei','微软雅黑',SimSun,sans-serif;font-size:14px;line-height:1.8;">您好！<br><br>附件为我司发出的往来及交易询证函，请查收并协助核对。<br>如函件信息无误，请在“信息证明无误”处签章；如有不符，请列明不符项目及具体内容。<br>完成后烦请将签章后的文件回传，谢谢！<br><br></div>"""
 m=month.strip() or f'{datetime.now().month}月份'
 if m.endswith('月'):m+='份'
 if not m.endswith('月份'):m+='月份'
 return f'【请查收】{c["name"]} - {m}对账单',"""<div style="font-family:'Microsoft YaHei','微软雅黑',SimSun,sans-serif;font-size:14px;line-height:1.8;">您好！<br><br>附件为本月对账单，请查收并尽快核对；如无误，请安排开票。<br><br>新电子发票要求：<br>1）仅需提供 OFD 和 XML 文件，PDF 文件无需提供；<br>2）发票命名：保留网上下载后的原文件名，并在文件名前增加贵公司简称；<br>3）截止收票日期：每月15日。<br><br></div>"""

class App:
 def __init__(self,root):
  self.root=root; root.title('Outlook 对账单 / 询证函邮件发送工具'); root.geometry('760x500')
  self.folder=tk.StringVar(); self.book=tk.StringVar(); self.month=tk.StringVar(value=f'{datetime.now().month}月份'); self.status=tk.StringVar(value='请选择附件文件夹和通讯录。')
  f=ttk.Frame(root,padding=18); f.pack(fill='both',expand=True); f.columnconfigure(1,weight=1); f.rowconfigure(5,weight=1)
  ttk.Label(f,text='Outlook 对账单 / 询证函邮件发送工具',font=('Microsoft YaHei UI',16,'bold')).grid(row=0,column=0,columnspan=3,sticky='w',pady=(0,16))
  self.row(f,1,'附件文件夹：',self.folder,self.pick_folder,'选择文件夹'); self.row(f,2,'供应商通讯录：',self.book,self.pick_book,'选择通讯录')
  ttk.Label(f,text='对账月份：').grid(row=3,column=0,sticky='e',pady=7); ttk.Entry(f,textvariable=self.month,width=16).grid(row=3,column=1,sticky='w',padx=8)
  ttk.Label(f,text='支持 PDF / XLS / XLSX / XLSM；通讯录支持“名称/收件人/抄送人”和“名称/邮箱/抄送”。\nPDF 自动按询证函生成，Excel 自动按对账单生成；仅打开预览，不自动发送。',foreground='#555').grid(row=4,column=0,columnspan=3,sticky='w',pady=8)
  box=ttk.LabelFrame(f,text='处理记录',padding=6); box.grid(row=5,column=0,columnspan=3,sticky='nsew'); box.rowconfigure(0,weight=1); box.columnconfigure(0,weight=1)
  self.log=tk.Text(box,state='disabled',wrap='word'); self.log.grid(row=0,column=0,sticky='nsew')
  ttk.Label(f,textvariable=self.status).grid(row=6,column=0,columnspan=2,sticky='w',pady=12); self.btn=ttk.Button(f,text='生成 Outlook 邮件预览',command=self.run); self.btn.grid(row=6,column=2,sticky='e')
 def row(self,f,r,label,var,cmd,btxt):
  ttk.Label(f,text=label).grid(row=r,column=0,sticky='e',pady=7); ttk.Entry(f,textvariable=var).grid(row=r,column=1,sticky='ew',padx=8); ttk.Button(f,text=btxt,command=cmd).grid(row=r,column=2)
 def pick_folder(self):
  x=filedialog.askdirectory(title='选择存放 PDF / Excel 附件的文件夹')
  if x:self.folder.set(x)
 def pick_book(self):
  x=filedialog.askopenfilename(title='选择供应商通讯录',filetypes=[('Excel 文件','*.xlsx *.xls *.xlsm')])
  if x:self.book.set(x)
 def write(self,x):
  self.log.config(state='normal'); self.log.insert('end',x+'\n'); self.log.see('end'); self.log.config(state='disabled'); self.root.update_idletasks()
 def run(self):
  folder=self.folder.get().strip(); book=self.book.get().strip()
  if not os.path.isdir(folder) or not os.path.isfile(book):return messagebox.showwarning('提示','请选择有效的附件文件夹和通讯录。')
  self.btn.config(state='disabled'); self.log.config(state='normal'); self.log.delete('1.0','end'); self.log.config(state='disabled'); ok=skip=fail=0
  try:
   contacts=read_contacts(book); self.write(f'已读取 {len(contacts)} 个供应商联系人。')
   files=sorted([p for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() in EXTS],key=lambda p:p.name.casefold())
   if not files:return messagebox.showwarning('提示','文件夹中没有 PDF 或 Excel 文件。')
   outlook=win32.Dispatch('Outlook.Application')
   for p in files:
    c=match(p.name,contacts)
    if not c:self.write(f'⚠ 未匹配：{p.name}'); skip+=1; continue
    if not c['to']:self.write(f'⚠ 无收件人：{c["name"]}'); skip+=1; continue
    try:
     subject,body=mail_text(c,p,self.month.get()); mail=outlook.CreateItem(0); mail.Display(); mail.To=c['to']
     if c['cc']:mail.CC=c['cc']
     mail.Subject=subject; mail.HTMLBody=body+mail.HTMLBody; mail.Attachments.Add(str(p.resolve())); ok+=1; self.write(f'✓ {p.name} → {c["to"]}'+(f'；抄送 {c["cc"]}' if c['cc'] else ''))
    except Exception as e:self.write(f'✗ {p.name}：{e}'); fail+=1
   self.status.set(f'完成：成功 {ok}，跳过 {skip}，失败 {fail}'); messagebox.showinfo('完成',f'已生成 {ok} 封邮件预览。\n跳过 {skip}，失败 {fail}。\n请检查后手动发送。')
  except Exception as e:self.write(traceback.format_exc()); messagebox.showerror('运行出错',str(e))
  finally:self.btn.config(state='normal')

if __name__=='__main__':
 root=tk.Tk(); App(root); root.mainloop()
