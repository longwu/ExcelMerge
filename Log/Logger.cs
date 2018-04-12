using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using Configuration;

namespace Log
{
    public class Logger : IDisposable
    {
        private static Logger _Logger = null;
        private BaseLoggerHandler _DefaultHandler = null;
        private ELogLevel _LogLevel = ELogLevel.None;

        public event Action<ELogLevel, string, string> AppendLog;

        public Logger(BaseLoggerHandler handler)
        {
            this._DefaultHandler = handler;
        }

        public Logger(ELogLevel level, BaseLoggerHandler handler)
        {
            this._LogLevel = level;
            this._DefaultHandler = handler;
        }

        public ELogLevel LogLevel
        {
            get { return this._LogLevel; }
            set { this._LogLevel = value; }
        }

        public static Logger Instance
        {
            get
            {
                if (_Logger == null)
                {
                    string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Const.AppName, ".log");
                    if (!Directory.Exists(logFilePath))
                        Directory.CreateDirectory(logFilePath);
                    logFilePath = Path.Combine(logFilePath, string.Format("{0}.log", DateTime.Now.ToString("yyyy-MM-dd")));
                    _Logger = new Logger(new BasicFileLoggerHandler(logFilePath));
                }
                return _Logger;
            }
        }

        public void Log(ELogLevel level, string tag, string message)
        {
#if DEBUG
            Debug.WriteLine(level.ToString() + " " + tag + " " + message);
#endif
            try
            {
                this.OnAppendLog(level, tag, message);
                if ((this._LogLevel & level) == level)
                {
                    if (this._DefaultHandler != null)
                    {
                        this._DefaultHandler.Log(level, tag, message);
                    }
                }
            }
            catch { }
        }

        private void OnAppendLog(ELogLevel level, string tag, string message)
        {
            if (this.AppendLog != null)
                this.AppendLog(level, tag, message);
        }

        public void Dispose()
        {
            if (this._DefaultHandler != null)
            {
                try
                {
                    this._DefaultHandler.Shutdown();
                }
                catch { }
            }
            this._DefaultHandler = null;
        }
    }
}
