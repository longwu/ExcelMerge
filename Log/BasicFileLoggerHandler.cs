using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Log
{
    /// <summary>
    /// 简单文件记录器。所有消息写到文本文件。
    /// </summary>
    public class BasicFileLoggerHandler : BaseLoggerHandler
    {
        const int WRITEFILE_BUFFER = 4096;

        private StreamWriter _Stream = null;
        private bool _Append = true;

        public BasicFileLoggerHandler(string filename)
        {
            if (!string.IsNullOrEmpty(filename))
            {
                FileMode fm;
                if (this._Append)
                    fm = FileMode.Append;
                else fm = FileMode.Create;

                FileStream fs = new FileStream(filename, fm, FileAccess.Write, FileShare.Read);

                this._Stream = new StreamWriter(fs, System.Text.Encoding.UTF8, WRITEFILE_BUFFER);
            }
        }

        /// <summary>
        /// 是否追加到日志文件尾部
        /// </summary>
        public bool Append
        {
            get { return this._Append; }
            set { this._Append = value; }
        }

        protected override void Log(LoggerMessage message)
        {
            if (this._Stream == null)
                return;
            try
            {
                this._Stream.Write(string.Format("[{0}][{1}:{2}]{3}\r\n", message.Time.ToString("yyyy-MM-dd HH:mm:ss.sss"), message.Level.ToString(), message.Tag, message.Message));
            }
            catch { }
        }

        protected override void OnShutdown()
        {
            if (this._Stream != null)
            {
                try
                {
                    this._Stream.Flush();
                    this._Stream.Close();
                }
                catch { }
            }
        }
    }
}
