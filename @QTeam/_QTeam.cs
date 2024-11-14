
using EasyPyFile;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Xml;
using static QTeam.QTLibrary;
using static System.Net.WebRequestMethods;
using Application = System.Windows.Forms.Application;
using WinForms = System.Windows.Forms;

namespace QTeam
{
    [Guid(QTeamUI.TypeGuid)]
    [ProgId(QTeamUI.TypeProgId)]
    [ClassInterface(ClassInterfaceType.AutoDual)]    
    
    public class QTeamUI : IDTExtensibility2, IRibbonExtensibility
    {
        #region Constants & Fields
        public const string TypeGuid = "9d3caed1-2003-490b-8123-0a6150308b49";
        public const string TypeProgId = "QTeam.UI";
        public const string ClsIdKeyName = @"CLSID\{" + QTeamUI.TypeGuid + @"}\";
        public const string ExcelAddInKeyName = @"Software\Microsoft\Office\Excel\Addins\" + QTeamUI.TypeProgId;
        static Microsoft.Office.Interop.Excel.Application _excel;
        static Microsoft.Office.Interop.Excel._Worksheet _workSheet;
        private IRibbonUI _ribbon;

        #endregion

        #region COM Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            Assembly asm = t.Assembly;
            Version v = asm.GetName().Version;
            RegistryKey QTeamkey;
            QTeamkey = Registry.ClassesRoot.CreateSubKey(QTeamUI.ClsIdKeyName + "Programmable");
            QTeamkey.Close();
            QTeamkey = Registry.ClassesRoot.CreateSubKey(QTeamUI.ClsIdKeyName + "InprocServer32");
            QTeamkey.SetValue(string.Empty, Environment.SystemDirectory + @"\mscoree.dll");
            QTeamkey.Close();
            QTeamkey = Registry.LocalMachine.CreateSubKey(QTeamUI.ExcelAddInKeyName);
            QTeamkey.SetValue("Description", "Fubon QTeam Systems", RegistryValueKind.String);
            QTeamkey.SetValue("FriendlyName", "Fubon QTeam UserInterface Add-In COM", RegistryValueKind.String);
            QTeamkey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            QTeamkey.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
            QTeamkey.SetValue("ApplicationVersion", v.ToString());
            QTeamkey.Close();
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            Registry.ClassesRoot.DeleteSubKey(QTeamUI.ClsIdKeyName + "Programmable");
            Registry.ClassesRoot.DeleteSubKeyTree(QTeamUI.ClsIdKeyName + "InprocServer32");
            Registry.LocalMachine.DeleteSubKey(QTeamUI.ExcelAddInKeyName);
        }
        #endregion

        #region IDTExtensibility2
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {        
            _excel = (Microsoft.Office.Interop.Excel.Application)Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            
        }

        public void OnStartupComplete(ref Array custom)
        {
            
        }

        public void OnBeginShutdown(ref Array custom)
        {
            
        }
        #endregion

        #region IRibbonExtensibility
        public string GetCustomUI(string RibbonID)
        {
            try
            {
                return QTeam.Properties.Resources.QTRibbon;
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(
                $@"@QTeam操作About時發生Exception.\n\t
                {ex.Source} \n\t
                {ex.Message}",
                "@QTeam Add-In",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            return null;
        }

        public void OnRibbon(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;    
        }
        
        string selectItem = null;
        public void OnQuerySelectionChanged(IRibbonControl control, string dropDownId, int selectIndex)
        {            
            selectItem = dropDownId;
            switch (selectItem)
            {
                case "匯率":                    
                    DisplayDataInsertExcel(selectItem, GetCurrency());
                    break;
                case "上櫃交易量":
                    DateTime dt = DateTime.Now;
                    CultureInfo culture = new CultureInfo("zh-TW");
                    culture.DateTimeFormat.Calendar = new TaiwanCalendar();
                    var d = dt.ToString("yyy/MM", culture);
                    DisplayDataInsertExcel(selectItem, DoWorkWebClient("OTC", d,
                        $"https://www.tpex.org.tw/web/stock/aftertrading/trading_amount/amt_rank_result.php?l=zh-tw&t=M&d={dt.ToString("yyy/MM/dd", culture)}"), d);
                    break;
                default:
                    break;
            }
        }

        string keys = null;
        public void OnKeysChange(IRibbonControl control, string text)
        {
            keys = text;
        }

        public string GetKeysText(IRibbonControl control)
        {
            return keys;
        }

        public void OnSearchButtonClick(IRibbonControl control)
        {     
            var keyWords = keys.Split(',');
            DoWorkMessage(keyWords);
            DisplayDataInsertExcel("新聞", ItemCache.InnerCache);
        }

        public void OnAboutButtonClick(IRibbonControl contorl)
        {
            try
            {                
                    WinForms.MessageBox.Show(
$@"@QTeam用於Excel自動交易及數據查詢功能, 
若發生任何系統問題, 請儘速聯絡程式交易科協助排除.

Functions:
WatchRoot- 取得文件根目錄.
CreateFile- 自動產生文件檔, Demo時使用.

Commands:
查詢數據-

",
            "@QTeam Add-In",
            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(
                $@"@QTeam操作About時發生Exception.\n\t
                {ex.Source} \n\t
                {ex.Message}",
                "@QTeam Add-In",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
        #endregion

        #region 交易量
        static private JArray DoWorkWebClient(string _type, string yyyMM, string _url)
        {
            try
            {
                switch (_type)
                {
                    case "OTC":
                        using (WebClient webClient = new WebClient())
                        {
                            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls;
                            webClient.Encoding = Encoding.UTF8;
                            var jsonStr = webClient.DownloadString(_url);
                            var jTokens = JsonConvert.DeserializeObject<JToken>(jsonStr);
                            var jArray = jTokens["tables"][0].Value<JArray>("data");
                            var date = jTokens["tables"][0].Value<string>("date");
                            if ( date == yyyMM)
                                return jArray;
                        }                        
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return null;
        }

        private void DisplayDataInsertExcel(string selectItem, JArray obj, string yyyMM)
        {
            try
            {
                _excel.Visible = true;
                _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_excel.ActiveSheet;
                _workSheet.Range[$"A1:Z{_workSheet.UsedRange.Rows.Count}"].Delete();
                //工作表名稱
                _workSheet.Name = selectItem;
                //建表頭
                _workSheet.Cells[1, "A"] = "標的代碼";
                _workSheet.Cells[1, "B"] = "標的名稱";
                _workSheet.Cells[1, "C"] = "月份";
                _workSheet.Cells[1, "D"] = "交易量";
                _workSheet.Range["A1:D1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DimGray);
                //加入資料
                var row = 1;
                if (obj.Any())
                {
                    foreach (var item in obj)
                    {
                        decimal amount;
                        decimal.TryParse(item[3].ToString(), out amount);
                        row++;
                        _workSheet.Cells[row, "A"] = item[1].ToString();
                        _workSheet.Cells[row, "B"] = item[2].ToString();
                        _workSheet.Cells[row, "C"] = yyyMM;
                        ((Range)_workSheet.Cells[row, "D"]).NumberFormat = "@";
                        _workSheet.Cells[row, "D"] = amount * 1000;
                    }
                }
                _workSheet.Columns[1].AutoFit();
                _workSheet.Columns[2].AutoFit();
                _workSheet.Columns[3].AutoFit();
                _workSheet.Columns[4].AutoFit();
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(
                $@"@QTeam操作About時發生Exception.\n\t
                {ex.Source} \n\t
                {ex.Message}",
                "@QTeam Add-In",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
        #endregion

        #region 新聞
        void DoWorkMessage(string[] keyWords)
        {
            try
            {
                ItemCache.InnerCache.Clear();
                Task.Run(() =>
                {
                    using (var webClient = new WebClient())
                    {
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                        webClient.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);
                        webClient.Encoding = Encoding.UTF8;
                        string page = webClient.DownloadString($"https://mops.twse.com.tw/mops/web/t05sr01_1");
                        if (page.Contains("查無資料") || page.Contains("資料庫中查無需求資料")) return;
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(page);
                        //var s = doc.DocumentNode.SelectSingleNode("//table[@class='hasBorder']");
                        var lists = doc.DocumentNode.SelectSingleNode("//table[@class='hasBorder']")
                                                .Descendants("tr")
                                                .Skip(1)
                                                .Where(tr => tr.Elements("td").Count() > 1)
                                                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim().Split(new char[] { '(', '-', ')', '<', '[', ']' })[0]).ToList())
                                                .ToList();
                        if (lists != null && lists.Any())
                        {
                            lists.ForEach(list =>
                            {
                                try
                                {
                                    if (keyWords.Any(word => list[4].Contains(word)))
                                    {
                                        string key = $"{list[0]}";
                                        if (ItemCache.GetValueItem(key) == null)
                                        {
                                            string value = string.Join("|", list.ToArray());
                                            ItemCache.Update(key, value);                                           
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw ex;
                                }
                            });
                        }
                    }
                }).Wait(3000);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        void DownloadUpdate()
        {
            string url = "https://mops.twse.com.tw/mops/web/ajax_index?encodeURIComponent=1&stp=0";
            WebClient wc = new WebClient();
            wc.DownloadStringCompleted += wc_DownloadStringCompleted;
            wc.DownloadStringAsync(new Uri(url));
        }

        private void wc_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            OnSearchButtonClick(null);
        }

        private void DisplayDataInsertExcel(string selectItem, ConcurrentDictionary<string, string> dict)
        {
            try
            {
                _excel.Visible = true;
                _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_excel.ActiveSheet;
                _workSheet.Range[$"A1:Z{_workSheet.UsedRange.Rows.Count}"].Delete();
                //工作表名稱
                _workSheet.Name = selectItem;
                //建表頭
                _workSheet.Cells[1, "A"] = "項目";
                _workSheet.Cells[1, "B"] = "主旨";
                _workSheet.Range["A1:B1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DimGray);
                //加入資料
                var row = 1;
                if (dict.Any())
                {
                    foreach (var item in dict)
                    {
                        row++;
                        _workSheet.Cells[row, "A"] = item.Key;
                        _workSheet.Cells[row, "B"] = item.Value;
                    }
                }
                _workSheet.Columns[1].AutoFit();
                _workSheet.Columns[2].AutoFit();
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(
                $@"@QTeam操作About時發生Exception.\n\t
                {ex.Source} \n\t
                {ex.Message}",
                "@QTeam Add-In",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
        #endregion

        #region 匯率
        public Dictionary<string, decimal> GetCurrency()
        {
            Dictionary<string, decimal> result = new Dictionary<string, decimal>();
            try
            {
                var task = CurrencyUri("https://tw.rter.info/capi.php");
                result = ToDictionary(JsonConvert.DeserializeObject<Dictionary<string, CurrencyLib>>(task));
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }

        private string VersionUri(string _url)
        {
            using (WebClient webClient = new WebClient())
            {
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls;
                return webClient.DownloadString(_url);
            }
        }

        private string CurrencyUri(string _url)
        {
            using (WebClient webClient = new WebClient())
            {
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls;
                return webClient.DownloadString(_url);
            }
        }

        private Dictionary<string, decimal> ToDictionary(Dictionary<string, CurrencyLib> _list)
        {
            string title = null;
            Dictionary<string, decimal> Currency = new Dictionary<string, decimal>();
            try
            {
                decimal twd = _list["USDTWD"].Exrate;
                foreach (var item in _list)
                {
                    if (item.Key.Contains("USD"))
                    {
                        title = item.Key.Substring(3, item.Key.Length - 3);
                        if (!Currency.ContainsKey(title) && !string.IsNullOrWhiteSpace(title))
                            Currency.Add(title, Math.Round(twd / item.Value.Exrate, 5));
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Currency;
        }

        private void DisplayDataInsertExcel(string selectItem, Dictionary<string, decimal> dict)
        {
            try
            {
                _excel.Visible = true;
                _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_excel.ActiveSheet;
                _workSheet.Range[$"A1:Z{_workSheet.UsedRange.Rows.Count}"].Delete();
                //工作表名稱
                _workSheet.Name = selectItem;
                //建表頭
                _workSheet.Cells[1, "A"] = "項目";
                _workSheet.Cells[1, "B"] = "匯價";
                _workSheet.Range["A1:B1"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DimGray);
                //加入資料
                var row = 1;
                if (dict.Any())
                {
                    foreach (var item in dict)
                    {
                        row++;
                        _workSheet.Cells[row, "A"] = item.Key;
                        ((Range)_workSheet.Cells[row, "B"]).NumberFormat = "@";
                        _workSheet.Cells[row, "B"] = item.Value;
                    }
                }
                _workSheet.Columns[1].AutoFit();
                _workSheet.Columns[2].AutoFit();
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show(
                $@"@QTeam操作About時發生Exception.\n\t
                {ex.Source} \n\t
                {ex.Message}",
                "@QTeam Add-In",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Excel Addins
        public string QTeamTest()
        {
            return "QTeam Test OK.";
        }

        public string WatchRoot()
        {
            try
            {               
                return $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        string SheetName()
        {
            try
            {
                _Worksheet workSheet = QTeam.QTeamUI._excel.ActiveSheet;
                return workSheet.Name;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// test create signal txt
        /// </summary>
        /// <param name="symbol">標的</param>
        /// <param name="price">價格</param>
        /// <param name="side">買賣(1,0,-1)</param>
        /// <param name="qty">數量</param>
        /// <param name="month">月份(yyyyMM)</param>
        /// <param name="strategy">策略名稱</param>
        /// <returns></returns>
        public string CreateFile( string symbol, 
            string price, 
            string side, string qty,
             string month = null, string strategy = null)
        {
            try
            {
                string strategyName = null;
                if (!string.IsNullOrWhiteSpace(strategy))
                {
                    strategyName = strategy;
                }
                else
                {
                    strategyName = SheetName();
                }
                var fileName = $"0_{strategyName}_{symbol}_{price}_{qty}_{side}_{month}___";
                using (FileStream fs = System.IO.File.Create($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\{fileName}.txt"))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes("");
                    fs.Write(info, 0, info.Length);
                }
                return "True";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string SortFile(string strategy, string symbol,
           string price, string qty, string side, string month)
        {
            try
            {
                string strategyName = null;
                if (!string.IsNullOrWhiteSpace(strategy))
                {
                    strategyName = strategy;
                }
                else
                {
                    strategyName = SheetName();
                }
                var fileName = $"0_{strategyName}_{symbol}_{price}_{qty}_{side}_{month}___";
                using (FileStream fs = System.IO.File.Create($@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\{fileName}.txt"))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes("");
                    fs.Write(info, 0, info.Length);
                }
                return "True";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
