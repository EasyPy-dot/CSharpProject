using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace mt5MonitorService
{
    internal class IniFile
    {
        string Path;
        string EXE = Assembly.GetExecutingAssembly().GetName().Name;

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

        public IniFile(string IniPath = null)
        {
            Path = new FileInfo(IniPath ?? EXE + ".ini").FullName;
        }

        public string Read(string Key, string Section = null)
        {
            var RetVal = new StringBuilder(255);
            GetPrivateProfileString(Section ?? EXE, Key, "", RetVal, 255, Path);
            return RetVal.ToString();
        }

        public void Write(string Key, string Value, string Section = null)
        {
            WritePrivateProfileString(Section ?? EXE, Key, Value, Path);
        }
    }

    internal class MonitorFilePath
    {
        public List<string> filePath { get; set; } = new List<string>();
    }

    internal enum OIType
    {
        JsonType, IniType
    }

    static internal class CollectionExtensions
    {
        /// <summary>
        /// take blockingcollection and enqueue concurrentqueue
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <param name="cancellationToken"></param>
        /// <param name="queue"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">queue is a null refference</exception>
        static public IEnumerable<T> TakeBuffer<T>(this BlockingCollection<T> collection,
            CancellationToken cancellationToken, ConcurrentQueue<T> queue)
        {
            if (collection.TryTake(out T item, 1, cancellationToken))
                queue.Enqueue(item);
            return queue;
        }
    }
}
