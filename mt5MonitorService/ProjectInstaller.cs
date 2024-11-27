using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace mt5MonitorService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
            this.AfterInstall += new InstallEventHandler(ProjectInstaller_AfterInstall);
        }

        private void ProjectInstaller_AfterInstall(object sender, InstallEventArgs e)
        {
            //設定etc/hosts
            if (!IsCheckedMQL5())
            {
                using (StreamWriter w = File.AppendText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"drivers\etc\hosts")))
                {
                    w.WriteLine("127.0.0.1  www.mql5.com");
                }
            }
            //安裝其它相關檔案...
        }

        private bool IsCheckedMQL5()
        {
            using (StreamReader r = new StreamReader(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"drivers\etc\hosts")))
            {
                while (r.Peek() >= 0)
                {
                    if (r.ReadLine().Contains("www.mql5.com"))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
