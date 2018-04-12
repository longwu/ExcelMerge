using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Log
{
    public enum ELogLevel : byte
    {
        /// <summary>
        /// 不显示 
        /// </summary>
        None = 1,   
        /// <summary>
        /// 信息
        /// </summary>
        Info = 2,
        /// <summary>
        /// 调试
        /// </summary>
        Debug = 4,
        /// <summary>
        /// 警告
        /// </summary>
        War = 8,
        /// <summary>
        /// 错误
        /// </summary>
        Error = 16,
        /// <summary>
        /// 严重错误111
        /// </summary>
        Fatal = 32,
    }
}
