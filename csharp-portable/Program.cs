using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MailToolPortable
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    sealed class Contact
    {
        public string Name = "";
        public string To = "";
        public string Cc = "";
    }

    sealed class MainForm : Form
    {
        readonly TextBox folderBox = new TextBox();
        readonly TextBox contactBox = new TextBox();
        readonly TextBox monthBox = new TextBox();
        readonly TextBox logBox = new TextBox();
        readonly Button runButton = new Button();

        public MainForm()
        {
            Text = "Outlook 对账单 / 询证函邮件发送工具（便携版）";
            Width = 820; Height = 580; StartPosition = FormStartPosition.CenterScreen;
            Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);

            var title = new Label { Text = "Outlook 对账单 / 询证函邮件发送工具", Left = 22, Top = 18, Width = 600, Height = 30, Font = new System.Drawing.Font(Font.FontFamily, 16F, System.Drawing.FontStyle.Bold) };
            Controls.Add(title);

            AddRow("附件文件夹：", folderBox, 65, PickFolder);
            AddRow("供应商通讯录：", contactBox, 105, PickContact);
            Controls.Add(new Label { Text = "对账月份：", Left = 22, Top = 151, Width = 100 });
            monthBox.SetBounds(125, 146, 160, 28); monthBox.Text = DateTime.Now.Month + "月份"; Controls.Add(monthBox);

            Controls.Add(new Label { Text = "支持 PDF / XLS / XLSX / XLSM；通讯录支持“名称/收件人/抄送人”和“名称/邮箱/抄送”。\r\nPDF 自动按询证函生成，Excel 自动按对账单生成；只打开 Outlook 预览，不自动发送。", Left = 22, Top = 188, Width = 750, Height = 46 });

            logBox.SetBounds(22, 242, 760, 245); logBox.Multiline = true; logBox.ScrollBars = ScrollBars.Vertical; logBox.ReadOnly = true; Controls.Add(logBox);
            runButton.Text = "生成 Outlook 邮件预览"; runButton.SetBounds(585, 500, 197, 34); runButton.Click += (_, __) => Run(); Controls.Add(runButton);
        }

        void AddRow(string label, TextBox box, int top, Action picker)
        {
            Controls.Add(new Label { Text = label, Left = 22, Top = top + 6, Width = 100 });
            box.SetBounds(125, top, 545, 28); Controls.Add(box);
            var b = new Button { Text = "选择", Left = 683, Top = top - 1, Width = 99, Height = 30 };
            b.Click += (_, __) => picker(); Controls.Add(b);
        }

        void PickFolder()
        {
            using (var d = new FolderBrowserDialog()) if (d.ShowDialog() == DialogResult.OK) folderBox.Text = d.SelectedPath;
        }

        void PickContact()
        {
            using (var d = new OpenFileDialog { Filter = "Excel 文件|*.xlsx;*.xls;*.xlsm" }) if (d.ShowDialog() == DialogResult.OK) contactBox.Text = d.FileName;
        }

        static string S(object v) => v == null ? "" : Convert.ToString(v)?.Trim() ?? "";
        static string Norm(string s) => Regex.Replace((s ?? "").Replace('（','(').Replace('）',')'), @"[\s　]+", "").TrimEnd('.', '_', '-', '—', '–').ToLowerInvariant();

        Dictionary<string, Contact> ReadContacts(string path)
        {
            dynamic excel = null, book = null, sheet = null, used = null;
            var result = new Dictionary<string, Contact>();
            try
            {
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false; excel.DisplayAlerts = false;
                book = excel.Workbooks.Open(path, ReadOnly: true); sheet = book.Worksheets[1]; used = sheet.UsedRange;
                int rows = used.Rows.Count, cols = used.Columns.Count;
                int nameCol = 0, toCol = 0, ccCol = 0;
                for (int c = 1; c <= cols; c++)
                {
                    string h = Norm(S(used.Cells[1, c].Value2));
                    if (new[] { "名称", "供应商名称", "公司名称" }.Contains(h)) nameCol = c;
                    if (new[] { "收件人", "邮箱", "收件邮箱", "电子邮箱", "email", "e-mail" }.Contains(h)) toCol = c;
                    if (new[] { "抄送人", "抄送", "抄送邮箱", "cc" }.Contains(h)) ccCol = c;
                }
                if (nameCol == 0 || toCol == 0) throw new Exception("通讯录必须包含“名称”和“收件人”列；也兼容“名称/邮箱/抄送”。");
                for (int r = 2; r <= rows; r++)
                {
                    string name = S(used.Cells[r, nameCol].Value2), to = S(used.Cells[r, toCol].Value2);
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var ccs = new List<string>();
                    if (ccCol > 0) for (int c = ccCol; c <= cols; c++) { var v = S(used.Cells[r, c].Value2); if (v.Contains("@")) ccs.Add(v); }
                    result[Norm(name)] = new Contact { Name = name, To = to, Cc = string.Join(";", ccs.Distinct()) };
                }
                return result;
            }
            finally
            {
                try { if (book != null) book.Close(false); } catch { }
                try { if (excel != null) excel.Quit(); } catch { }
                Release(used); Release(sheet); Release(book); Release(excel);
            }
        }

        static void Release(object o) { try { if (o != null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); } catch { } }

        Contact Match(string file, Dictionary<string, Contact> contacts)
        {
            string stem = Path.GetFileNameWithoutExtension(file).Trim();
            var options = new List<string> { stem };
            foreach (char x in new[] { '_', '-', '—', '–' }) { int i = stem.IndexOf(x); if (i > 0) options.Add(stem.Substring(0, i).Trim()); }
            options.Add(Regex.Replace(stem, @"[\s_-]*(?:\d{4}年)?\d{1,2}月(?:份)?对账单$", "").Trim());
            options.Add(Regex.Replace(stem, @"[\s_-]*(?:往来及交易)?询证函$", "").Trim());
            foreach (var x in options) if (contacts.TryGetValue(Norm(x), out var c)) return c;
            string ns = Norm(stem); return contacts.Where(k => ns.Contains(k.Key)).OrderByDescending(k => k.Key.Length).Select(k => k.Value).FirstOrDefault();
        }

        void Run()
        {
            if (!Directory.Exists(folderBox.Text) || !File.Exists(contactBox.Text)) { MessageBox.Show("请选择有效的附件文件夹和通讯录。", "提示"); return; }
            runButton.Enabled = false; logBox.Clear(); int ok = 0, skip = 0, fail = 0;
            try
            {
                var contacts = ReadContacts(contactBox.Text); Log($"已读取 {contacts.Count} 个供应商联系人。");
                var files = Directory.GetFiles(folderBox.Text).Where(f => new[] { ".pdf", ".xls", ".xlsx", ".xlsm" }.Contains(Path.GetExtension(f).ToLowerInvariant())).OrderBy(Path.GetFileName).ToList();
                if (files.Count == 0) { MessageBox.Show("文件夹中没有 PDF 或 Excel 文件。", "提示"); return; }
                dynamic outlook = Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application"));
                foreach (var f in files)
                {
                    var c = Match(f, contacts);
                    if (c == null) { Log("未匹配：" + Path.GetFileName(f)); skip++; continue; }
                    if (string.IsNullOrWhiteSpace(c.To)) { Log("无收件人：" + c.Name); skip++; continue; }
                    try
                    {
                        dynamic mail = outlook.CreateItem(0); mail.Display(); mail.To = c.To; if (!string.IsNullOrWhiteSpace(c.Cc)) mail.CC = c.Cc;
                        bool pdf = Path.GetExtension(f).Equals(".pdf", StringComparison.OrdinalIgnoreCase);
                        string month = string.IsNullOrWhiteSpace(monthBox.Text) ? DateTime.Now.Month + "月份" : monthBox.Text.Trim();
                        if (month.EndsWith("月")) month += "份";
                        mail.Subject = pdf ? $"【请查收】{c.Name} - 往来及交易询证函" : $"【请查收】{c.Name} - {month}对账单";
                        string body = pdf ? "您好！<br><br>附件为我司发出的往来及交易询证函，请查收并协助核对。<br>如函件信息无误，请在“信息证明无误”处签章；如有不符，请列明不符项目及具体内容。<br>完成后烦请将签章后的文件回传，谢谢！<br><br>" : "您好！<br><br>附件为本月对账单，请查收并尽快核对；如无误，请安排开票。<br><br>新电子发票要求：<br>1）仅需提供 OFD 和 XML 文件，PDF 文件无需提供；<br>2）发票命名：保留网上下载后的原文件名，并在文件名前增加贵公司简称；<br>3）截止收票日期：每月15日。<br><br>";
                        mail.HTMLBody = "<div style=\"font-family:Microsoft YaHei;font-size:14px;line-height:1.8\">" + body + "</div>" + mail.HTMLBody;
                        mail.Attachments.Add(Path.GetFullPath(f)); ok++; Log($"成功：{Path.GetFileName(f)} → {c.To}" + (string.IsNullOrWhiteSpace(c.Cc) ? "" : "；抄送 " + c.Cc));
                        Release(mail);
                    }
                    catch (Exception ex) { fail++; Log("失败：" + Path.GetFileName(f) + "；" + ex.Message); }
                }
                Release(outlook); MessageBox.Show($"已生成 {ok} 封邮件预览。\n跳过 {skip}，失败 {fail}。\n请检查后手动发送。", "完成");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "运行出错"); Log(ex.ToString()); }
            finally { runButton.Enabled = true; }
        }

        void Log(string s) { logBox.AppendText(s + Environment.NewLine); Application.DoEvents(); }
    }
}
