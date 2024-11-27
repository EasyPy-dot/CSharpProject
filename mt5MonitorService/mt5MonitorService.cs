
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace mt5MonitorService
{
    public partial class mt5MonitorService : ServiceBase
    {
        public mt5MonitorService()
        {
            InitializeComponent();
        }
        private Logger nLog = NLog.LogManager.GetCurrentClassLogger();
        private FileSystemWatcher watcher;
        private bool isChecked;
        private string userName;
        private bool isAdministrator = true;
        private string[] dirs;
        private string jsonString;
        private ManagementEventWatcher startWatch = null;
        private ManagementEventWatcher stopWatch = null;
        private MonitorFile monitorFile;
        private List<MonitorFile> fileEvent;
        CancellationTokenSource cancellationTokenSource;
        BlockingCollection<BufferObject> blockingCollection;
        ConcurrentQueue<BufferObject> concurrentQueue;

        /// <summary>
        /// 服務開始時
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            try
            {
                nLog.Info("監控服務重啟!");
                //開機時
                isChecked = true;
                monitorFile = new MonitorFile(null);
                monitorFile.OnKillProcess += MonitorFile_OnKillProcess;
                monitorFile.OnWriteFile += MonitorFile_OnWriteFile;
                monitorFile.OnReadFile += MonitorFile_OnReadFile;
                cancellationTokenSource = new CancellationTokenSource();
                blockingCollection = new BlockingCollection<BufferObject>();
                concurrentQueue = new ConcurrentQueue<BufferObject>();
                fileEvent = new List<MonitorFile>();
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
                    ManagementObjectCollection collection = searcher.Get();
                    userName = (string)collection.Cast<ManagementBaseObject>().First()["UserName"];
                }
                catch
                {
                    isAdministrator = false;
                    nLog.Info("無法使用管理者權限!!!");
                    userName = (System.Security.Principal.WindowsIdentity.GetCurrent().Name);
                }
                //沒有installer service時, 設置hosts
                SethostsParameter();
                StartMonitorFile();
                if (isAdministrator)
                {
                    nLog.Info("可使用管理者權限!!!");
                    StartMonitorProcess();
                }
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        protected void SethostsParameter()
        {
            try
            {
                //設定etc/hosts
                if (!IsCheckedMQL5())
                    using (StreamWriter w = File.AppendText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"drivers\etc\hosts")))
                        w.WriteLine("127.0.0.1  www.mql5.com");
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        private bool IsCheckedMQL5()
        {
            using (StreamReader r = new StreamReader(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"drivers\etc\hosts")))
            {
                while (r.Peek() >= 0)
                {
                    if (r.ReadLine().Contains("www.mql5.com"))
                        return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 監控檔案程序
        /// </summary>
        protected void StartMonitorFile()
        {
            try
            {
                nLog.Info("開始監控服務!");
                nLog.Info($"{AppDomain.CurrentDomain.BaseDirectory}");
                monitorFile.ReadFile(OIType.JsonType, $@"{AppDomain.CurrentDomain.BaseDirectory}\path.json");
                nLog.Info($"{jsonString}");
                MonitorFilePath m = JsonConvert.DeserializeObject<MonitorFilePath>(jsonString);
                if (m == null || m.filePath == null || m.filePath.Count == 0)
                {
                    //取得AppData\Roaming路徑
                    var appPath = GetFilePath();
                    while (!Directory.Exists($@"{appPath}\MetaQuotes\Terminal"))
                        Thread.Sleep(5000);
                    dirs = Directory.GetDirectories($@"{appPath}\MetaQuotes\Terminal", "*.*", SearchOption.AllDirectories);
                    while (!dirs.Any(d => d.EndsWith("config")))
                    {
                        Thread.Sleep(1000);
                        dirs = Directory.GetDirectories($@"{appPath}\MetaQuotes\Terminal", "*.*", SearchOption.AllDirectories);
                    }
                    monitorFile.KillProcess("terminal64");
                    Thread.Sleep(2000);
                    dirs = dirs.Where(d => d.EndsWith("config")).ToArray();
                    ChangeFile(dirs);
                }
                if (m != null && m.filePath != null && m.filePath.Count > 0)
                    dirs = m.filePath.ToArray();
                //file watcher監控資料
                foreach (var dir in dirs)
                    monitorFile.ReadFile(OIType.IniType, dir);
            }
            catch( Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// Process Watcher
        /// </summary>
        protected void StartMonitorProcess()
        {
            if (startWatch == null && stopWatch == null)
                StartMonitorManagementEvent();
        }

        /// <summary>
        /// Start Management Event Watcher
        /// </summary>
        protected void StartMonitorManagementEvent()
        {
            try
            {
                Task.Run(() =>
                {
                    try
                    {
                        if (startWatch == null)
                        {
                            startWatch = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStartTrace"));
                            startWatch.EventArrived += StartWatch_EventArrived;
                        }
                        startWatch.Start();
                        if (stopWatch == null)
                        {
                            stopWatch = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStopTrace"));
                            stopWatch.EventArrived += StopWatch_EventArrived;
                        }
                        stopWatch.Start();
                    }
                    catch (Exception ex)
                    {
                        nLog.Error($"{ex.Message} \n\r {ex.Source}");
                    }
                });
            }
            catch ( Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// writer file
        /// </summary>
        /// <param name="arg1">file type</param>
        /// <param name="arg2">checked parameters is exist</param>
        /// <param name="arg3">file path</param>
        private void MonitorFile_OnWriteFile(OIType arg1, int arg2, string arg3)
        {
            try
            {
                switch (arg1)
                {
                    case OIType.JsonType:
                        System.IO.File.WriteAllText($@"{AppDomain.CurrentDomain.BaseDirectory}\path.json", arg3);
                        break;
                    case OIType.IniType:
                        var ini = new IniFile($@"{ arg3 }\common.ini");
                        if (arg2 == 0)//writting Services value
                        {
                            ini.Write("Services", "4294967168", "Common");
                        }
                        else//writting Common tag
                        {
                            ini.Write("Login", "0", "Common");
                            ini.Write("Server", "", "Common");
                            ini.Write("ProxyEnable", "0", "Common");
                            ini.Write("ProxyType", "0", "Common");
                            ini.Write("ProxyAddress", "", "Common");
                            ini.Write("ProxyAuth", "", "Common");
                            ini.Write("CertInstall", "0", "Common");
                            ini.Write("NewsEnable", "1", "Common");
                            ini.Write("Services", "4294967168", "Common");
                            ini.Write("NewsLanguages", "", "Common");
                            ini.Write("Source", "download.mql5.com", "Common");
                            ini.Write("MQL5Login", "", "Common");
                            ini.Write("MQL5Password", "", "Common");
                        }
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// reader file
        /// </summary>
        /// <param name="arg1">file type</param>
        /// <param name="arg2">file path</param>
        private void MonitorFile_OnReadFile(OIType arg1, string arg2)
        {
            try
            {
                switch (arg1)
                {
                    case OIType.JsonType:
                        using (StreamReader r = new StreamReader(arg2))
                            jsonString = r.ReadToEnd();
                        break;
                    case OIType.IniType:
                        //加入監控路徑
                        fileEvent.Add(new MonitorFile(arg2));
                        fileEvent[fileEvent.Count - 1].OnAddFile += MonitorFile_OnChangedFile;
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// kill process
        /// </summary>
        /// <param name="obj">process name</param>
        private void MonitorFile_OnKillProcess(string obj)
        {
            try
            {
                var process = Process.GetProcessesByName(obj);
                if (process != null && process.Length > 0)
                {
                    foreach (var p in process)
                    {
                        if (!p.HasExited)
                        {
                            nLog.Info($"Kill >【 { p.Id} | { p.ProcessName} 】 {DateTime.Now}");
                            p.Kill();
                            p.WaitForExit();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// 取得AppData\Roaming路徑
        /// </summary>
        /// <returns></returns>
        private string GetFilePath()
        {
            try
            {
                string[] users = userName.Split('\\');
                userName = users.Length > 1 ? users[1] : users[0];
                //因為必需使用windows systems32啟動, 故Enviroment無法取得正確的使用者資料夾
                string appPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                appPath = appPath.Replace("WINDOWS", "Users");
                appPath = appPath.Replace("system32", userName);
                appPath = appPath.Replace(@"config\systemprofile\", "");
                return appPath;
            }
            catch(Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
            return null;
        }

        private void MonitorFile_OnChangedFile(object arg1, FileSystemEventArgs arg2)
        {
            try
            {
                nLog.Info("檔案變更: " + arg2.FullPath + " " + arg2.ChangeType);
                //生產者/消費者模式
                ProducerBuffer((FileSystemWatcher)arg1, arg2);
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// producer
        /// </summary>
        /// <param name="_watcher"></param>
        private void ProducerBuffer(FileSystemWatcher _watcher, FileSystemEventArgs _args)
        {
            if (blockingCollection.TryAdd(new BufferObject()
            {
                bufferData = _watcher,
                eventArgs = _args
            }))
            blockingCollection.TakeBuffer(cancellationTokenSource.Token, concurrentQueue);
            _watcher.Changed -= MonitorFile_OnChangedFile;
            ThreadPool.QueueUserWorkItem(new WaitCallback(ConsumerBuffer));
        }

        /// <summary>
        /// consumer
        /// </summary>
        /// <param name="_obj"></param>
        internal void ConsumerBuffer(object _obj)
        {
            while (concurrentQueue.TryDequeue(out BufferObject obj))
            {
                try
                {
                    nLog.Info("檔案變更: " + obj.eventArgs.FullPath + " " + obj.eventArgs.ChangeType);
                    System.Timers.Timer t = new System.Timers.Timer();
                    watcher = obj.bufferData;
                    t.Interval = 3000;
                    t.Elapsed += T_Elapsed;
                    t.Start();
                }
                catch (Exception ex)
                {
                    nLog.Error($"{ex.Message} \n\r {ex.Source}");
                }
            }
        }

        /// <summary>
        /// 初次安裝在結束pid後修改common.ini
        /// </summary>
        /// <param name="dirs"></param>
        protected void ChangeFile(string[] dirs)
        {
            try
            {
                MonitorFilePath file = new MonitorFilePath();
                foreach (var dir in dirs)
                {
                    monitorFile.WriteFile(OIType.IniType, dir, 1);
                    file.filePath.Add(dir);
                }
                monitorFile.WriteFile(OIType.JsonType, JsonConvert.SerializeObject(file));
            }
            catch (Exception ex)
            {

                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// 檔案異動時
        /// </summary>
        protected void ChangeFile()
        {
            try
            {
                if (dirs == null || dirs.Length == 0)
                {
                    nLog.Warn("dirs為空值, 請檢查程序.");
                }
                else
                {
                    //修改Services值, 讓欲啟用mql5服務的使用者無法啟用.
                    nLog.Info("修改Services值!!!");
                    foreach (var dir in dirs)
                        monitorFile.WriteFile(OIType.IniType, dir);
                }
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }       

        /// <summary>
        /// sleep time
        /// </summary>
        /// <param name="sender"></param>
        ///// <param name="e"></param>
        protected void T_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                ((System.Timers.Timer)sender).Stop();
                if (!isAdministrator)
                {
                    Process[] processesArray = Process.GetProcessesByName("terminal64");
                    foreach (Process process in processesArray)
                    {
                        if (process.ProcessName.ToLower() == "terminal64")
                        {
                            if (!process.HasExited)
                            {
                                process.EnableRaisingEvents = true;
                                process.Exited += Process_Exited;
                            }
                        }
                    }
                }
                watcher.Changed += MonitorFile_OnChangedFile;
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// 進程關閉
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Process_Exited(object sender, EventArgs e)
        {
            try
            {
                //重新檢查監視路徑
                ReCheckedPath();
                Thread.Sleep(1000);
                Process p = (Process)sender;
                ChangeFile();
                nLog.Info($"關閉{p.ProcessName}.exe");
            }
            catch(Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// stop watch process event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void StopWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            string name = e.NewEvent.Properties["ProcessName"].Value.ToString();
            if (name.Contains("terminal64"))
            {
                string pid = e.NewEvent.Properties["ProcessID"].Value.ToString();
                nLog.Info("修改Services值!!!");
                foreach (var dir in dirs)
                    monitorFile.WriteFile(OIType.IniType, dir);
                nLog.Info($"Stop >【 { pid } | { name } 】{ DateTime.Now }");
            }
        }

        /// <summary>
        /// start watch process event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void StartWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            string name = e.NewEvent.Properties["ProcessName"].Value.ToString();
            if (name.Contains("terminal64"))
            {
                //重新檢查監視路徑
                ReCheckedPath();
                //記錄pid
                string pid = e.NewEvent.Properties["ProcessID"].Value.ToString();
                nLog.Info($"Start >【 { pid} | { name} 】 {DateTime.Now}");
            }
        }

        /// <summary>
        /// 檢查監視路徑
        /// </summary>
        private void ReCheckedPath()
        {
            try
            {                
                //重新檢查監控路徑陣列
                var appPath = GetFilePath();
                dirs = Directory.GetDirectories($@"{appPath}\MetaQuotes\Terminal", "*.*", SearchOption.AllDirectories);
                dirs = dirs.Where(d => d.EndsWith("config")).ToArray();
                //已紀錄的路徑
                monitorFile.ReadFile(OIType.JsonType, $@"{AppDomain.CurrentDomain.BaseDirectory}\path.json");
                nLog.Info($"{jsonString}");
                MonitorFilePath m = JsonConvert.DeserializeObject<MonitorFilePath>(jsonString);
                //路徑陣列異動時
                if (dirs.Length > m.filePath.Count)
                {
                    try
                    {
                        fileEvent.Clear();
                    }
                    catch (Exception ex)
                    {
                        nLog.Error($"{ex.Message} \n\r {ex.Source}");
                    }
                    monitorFile.KillProcess("terminal64");
                    //file watcher監控資料
                    MonitorFilePath file = new MonitorFilePath();
                    foreach (var dir in dirs)
                    {
                        monitorFile.ReadFile(OIType.IniType, dir);
                        file.filePath.Add(dir);
                    }
                    monitorFile.WriteFile(OIType.JsonType, JsonConvert.SerializeObject(file));                    
                }
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }

        /// <summary>
        /// stop management watcher
        /// </summary>
        protected void StopMonitorManagementEvent()
        {
            startWatch.EventArrived -= StartWatch_EventArrived;
            stopWatch.EventArrived -= StopWatch_EventArrived;
            startWatch = null;
            stopWatch = null;
        }

        /// <summary>
        /// 服務停止時
        /// </summary>
        protected override void OnStop()
        {
            try
            {
                if (isChecked)
                {
                    try
                    {
                        foreach (var fevent in fileEvent)
                            fevent.OnAddFile -= MonitorFile_OnChangedFile;
                        fileEvent.Clear();
                    }
                    catch(Exception ex)
                    {
                        nLog.Error($"{ex.Message} \n\r {ex.Source}");
                    }
                    StopMonitorManagementEvent();
                    nLog.Warn("監控檔案及進程服務中止!");
                }
            }
            catch (Exception ex)
            {
                nLog.Error($"{ex.Message} \n\r {ex.Source}");
            }
        }
    }
}
