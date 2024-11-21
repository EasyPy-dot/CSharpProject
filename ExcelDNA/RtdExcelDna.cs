using ExcelDna.Integration;
using System;
using System.Threading;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Collections.Generic;

namespace SampleExcelDna
{
    public class RtdExcelDna
    {
        static readonly Dictionary<string, Subject<_x>> _tData =
            new Dictionary<string, Subject<_x>>();

        // this gets called by the server if there is a new value
        static public void SendValue(_x x)
        {
            if (!_tData.ContainsKey(x.name))
                _tData[x.name] = new Subject<_x>();
            Subject<_x> codeData = _tData[x.name];
            codeData.OnNext(x);
        }

        static IObservable<_x> GetObservable(string topic)
        {
            Subject<_x> subject;
            if (_tData.TryGetValue(topic, out subject))
            {
                return subject;
            }
            // Maybe better to define some kind of 'invalid' TickInfo...?
            return Observable.Return<_x>(null);
        }

        // this calls the wrapper to actually put it all together
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object QTRtdDna(string topic, string field)
        {
            string functionName = nameof(QTRtdDna);
            object[] paramInfo = new object[] { topic, field };
            return RxExcel.Observe(functionName, paramInfo, () => GetObservable(topic).Select<_x, object>(
                it =>
                {
                    if (it == null)
                    {
                        return "### 失敗的Topic !!!";
                    }
                    switch (field)
                    {
                        case "TIME":
                            return it.xTime;
                        case "DATE":
                            return it.xDate;
                        default:
                            return "### 失敗的Topic !!!";
                    }
                }));
        }
    }

    public static class RxExcel
    {
        public static IExcelObservable ToExcelObservable<T>(this IObservable<T> observable, object[] parameters)
        {
            return new ExcelObservable<T>(observable, parameters);
        }

        public static object Observe<T>(string functionName, object[] parameters, Func<IObservable<T>> observableSource)
        {
            return ExcelAsyncUtil.Observe(functionName, parameters, () => observableSource().ToExcelObservable(parameters));
        }
    }

    class ExcelObservable<T>: IExcelObservable
    {
        readonly IObservable<T> _observable;
        static Timer _timer;
        static object[] _params;
        static _x x;

        public ExcelObservable(IObservable<T> observable, object[] parameters)
        {
            _params = parameters;
            _observable = observable;
             x = new _x() { name = _params[0].ToString() };
            _timer = new Timer(timer_tick, null, 0, 1000);
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            observer.OnNext($"Subscribe: {_params[0]}.{_params[1]}");
            return _observable.Subscribe(value => observer.OnNext(value), observer.OnError, observer.OnCompleted);            
        }

        static void timer_tick(object _)
        {
            x.xDate = DateTime.Now.ToString("yyyy-MM-dd");
            x.xTime = DateTime.Now.ToString("HH:mm:ss.fff");
            RtdExcelDna.SendValue(x);
        }
    }

    //建立物件
    public class _x
    {
        public string name { get; set; }
        public string xTime { get; set; } = DateTime.Now.ToString("HH:mm:ss.fff");
        public string xDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
    }
}
