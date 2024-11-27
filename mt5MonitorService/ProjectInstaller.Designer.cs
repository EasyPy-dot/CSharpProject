using System;
using System.Collections;
using System.IO;

namespace mt5MonitorService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        protected override void OnCommitted(IDictionary savedState)
        {
            base.OnCommitted(savedState);
            //啟動windows services
            System.ServiceProcess.ServiceController sc = new System.ServiceProcess.ServiceController(serviceInstaller1.ServiceName);
            if (sc.Status == System.ServiceProcess.ServiceControllerStatus.Stopped)
                sc.Start();
            ////設定etc/hosts
            //using (StreamWriter w = File.AppendText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"drivers\etc\hosts")))
            //{
            //    w.WriteLine("127.0.0.1  www.mql5.com");
            //}
            ////安裝fubonfutures5setup
            //System.Diagnostics.Process.Start($@"{Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)}\fubonfutures5setup.exe");
        }

        #region 元件設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.serviceProcessInstaller1 = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstaller1 = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstaller1
            // 
            this.serviceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.serviceProcessInstaller1.Password = null;
            this.serviceProcessInstaller1.Username = null;
            // 
            // serviceInstaller1
            // 
            this.serviceInstaller1.Description = "monitor mt5";
            this.serviceInstaller1.DisplayName = "mt5MonitorService";
            this.serviceInstaller1.ServiceName = "mt5MonitorService";
            this.serviceInstaller1.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstaller1,
            this.serviceInstaller1});

        }

        #endregion

        protected System.ServiceProcess.ServiceProcessInstaller serviceProcessInstaller1;
        protected System.ServiceProcess.ServiceInstaller serviceInstaller1;
    }
}