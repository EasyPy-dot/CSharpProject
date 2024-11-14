using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QTeam
{
    internal class QTLibrary
    {
        [Serializable]
        public class CurrencyLib
        {
            public string UTC { get; set; }
            public decimal Exrate { get; set; }
        }
    }

    internal class ItemCache : Dictionary<string, string>
    {
        static internal ConcurrentDictionary<string, string> InnerCache { get; set; } = new ConcurrentDictionary<string, string>();
        static internal void Update(string _key, string _value)
        {
            try
            {
                InnerCache.TryAdd(_key, _value);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        static internal string GetValueItem(string _key)
        {
            try
            {
                if (InnerCache.TryGetValue(_key, out string value))
                    return value;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return null;
        }
    }
}
