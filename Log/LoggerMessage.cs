using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Log
{
    /// <summary>
    /// 一个日志消息
    /// </summary>
    public class LoggerMessage
    {
        public ELogLevel Level { get; set; }
        public string Tag { get; set; }
        public string Message { get; set; }
        public DateTime Time { get; set; }
    }
}
