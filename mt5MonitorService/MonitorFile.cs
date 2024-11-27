using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mt5MonitorService
{
    internal class MonitorFile: IDisposable
    {
        private bool disposedValue;
        //internal event Action<string[]> OnAddFilePath;
        internal event Action<string> OnKillProcess;
        internal event Action<OIType, int, string> OnWriteFile;
        internal event Action<OIType, string> OnReadFile;        
        internal event Action<object, FileSystemEventArgs> OnAddFile;
        private FileSystemWatcher fwatcher;
        public FileSystemWatcher watcher
        {
            get { return fwatcher; }
            set
            {
                if (fwatcher != null)
                {
                    fwatcher.Changed -= Fwatcher_Changed;
                }
                fwatcher = value;
                if (fwatcher != null)
                {
                    fwatcher.Changed += Fwatcher_Changed;
                }
            }
        }

        private void Fwatcher_Changed(object sender, FileSystemEventArgs e)
        {
            this.OnAddFile?.Invoke(sender, e);
        }

        public MonitorFile(string dir)
        {
            if (dir == null)
                return;
            watcher = new FileSystemWatcher
            {
                // 設定要監看的資料夾, dir
                Path = dir,
                // 設定要監看的變更類型，這裡設定監看最後修改時間與修改檔名的變更事件
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName,
                // 設定要監看的檔案類型
                Filter = "common.ini",
                // 設定是否監看子資料夾
                IncludeSubdirectories = false,
                // 設定是否啟動元件，必須要設定為 true，否則監看事件是不會被觸發
                EnableRaisingEvents = true
            };
        }

        internal void KillProcess(string process)
        {
            this.OnKillProcess?.Invoke(process);
        }

        internal void WriteFile(OIType otype, string file, int allType = 0)
        {
            this.OnWriteFile?.Invoke(otype, allType, file);
        }

        internal void ReadFile(OIType otype, string file)
        {
            this.OnReadFile?.Invoke(otype, file);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: 處置受控狀態 (受控物件)
                    //this.OnAddFilePath = null;
                    this.OnKillProcess = null;
                    this.OnWriteFile = null;
                    this.OnReadFile = null;
                    this.OnAddFile = null;
                }

                // TODO: 釋出非受控資源 (非受控物件) 並覆寫完成項
                // TODO: 將大型欄位設為 Null
                disposedValue = true;
            }
        }

        // // TODO: 僅有當 'Dispose(bool disposing)' 具有會釋出非受控資源的程式碼時，才覆寫完成項
        // ~MonitorFile()
        // {
        //     // 請勿變更此程式碼。請將清除程式碼放入 'Dispose(bool disposing)' 方法
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // 請勿變更此程式碼。請將清除程式碼放入 'Dispose(bool disposing)' 方法
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        ~MonitorFile() => Dispose(false);
    }

    internal class BufferObject
    {
        internal FileSystemWatcher bufferData { get; set; }
        internal FileSystemEventArgs eventArgs { get; set; }
    }
}
