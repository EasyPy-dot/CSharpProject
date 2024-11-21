using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Xml.Linq;

namespace SampleRtd
{
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("RtdExcel")]
    [ComVisible(true)]
    public class RtdExcel : IRtdServer
    {
        static Dictionary<string, string> _xData = new Dictionary<string, string>();
        readonly Dictionary<int, string> _topics = new Dictionary<int, string>();
        private System.Threading.Timer _timer;
        private System.Timers.Timer m_Timer;
        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            try
            {
                m_Timer = new System.Timers.Timer();
                m_Timer.Interval = 1000;
                m_Timer.Elapsed += Timer_Elapsed;
                m_Timer.Enabled = true;
                m_Timer.AutoReset = true;
                m_Timer.Start();
                _timer = new System.Threading.Timer(delegate { CallbackObject.UpdateNotify(); }, null, 1000, 1000);
                return 1;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                foreach (var item in new string[] { "DATE", "TIME" })
                {
                    if (_xData.ContainsKey(item))
                        _xData[item] = item.Contains("DATE") ? GetDate() : GetTime();
                    else
                        _xData.Add(item, item.Contains("DATE") ? GetDate() : GetTime());
                }
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        public dynamic ConnectData(int TopicID, ref Array Strings, ref bool GetNewValues)
        {
            try
            {
                var param1 = (string)Strings.GetValue(0);
                //多個數組可組成唯一標識Topic, 如時日.Date, 時日.Time
                //if (Strings.Length > 1)
                //param1 += "." + (string)Strings.GetValue(1);
                GetNewValues = true;
                _topics[TopicID] = param1;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _topics[TopicID];
        }

        public Array RefreshData(ref int TopicCount)
        {
            object[,] results = new object[2, _topics.Count];
            try
            {
                TopicCount = 0;
                foreach (KeyValuePair<int, string> kvp in _topics)
                {
                    results[0, TopicCount] = kvp.Key;
                    var item = kvp.Value;
                    if (_xData.ContainsKey(item))
                    {
                        //二維數組中取得要刷新的資料至對應的儲存格
                        //if (item.Length > 1)
                        //{
                        //switch (item[1])
                        //{
                        //case "DATE":
                        //results[1, TopicCount] = _xData[item[0]].Date;
                        //break;
                        // default:
                        //results[1, TopicCount] = _xData[item[0]].Time;
                        //break;
                        //}
                        //}
                        results[1, TopicCount] = _xData[item];
                    }
                    TopicCount++;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return results;
        }

        public void DisconnectData(int TopicID)
        {
            try
            {
                _topics.Remove(TopicID);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public int Heartbeat()
        {
            return 1;
        }

        string GetDate()
        {
            return DateTime.Now.ToString("yyyy-mm-dd");
        }

        string GetTime()
        {
            return DateTime.Now.ToString("hh:mm:ss:fff");
        }

        public void ServerTerminate()
        {
            try
            {
                _timer.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region Register add in
        /// <summary>
        /// AddIn Register Function
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + t.GUID.ToString().ToUpper() + @"}\Programmable");//
                var key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(@"CLSID\{" + t.GUID.ToString().ToUpper() + @"}\InprocServer32", true);//
                if (key != null)
                    key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll", Microsoft.Win32.RegistryValueKind.String);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// AddIn Unregister Function
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + t.GUID.ToString().ToUpper() + @"}\Programmable");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
