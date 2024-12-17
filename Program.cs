using AutoUpdaterDotNET;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace AutoUpdaterTest
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            await Task.Run(() =>
            {
                AutoUpdateTools.StartUpdaterAsync("https://github.com/EasyPy-dot/CSharpProject/raw/refs/heads/autoupdater_sample/AutoUpdater.json");
                
            });
            await Task.Delay(15000);
        }
    }
}
