using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerge
{
    /// <summary>
    /// 循环引擎接口
    /// </summary>
    public interface ILoopEngine
    {
        /// <summary>
        /// 循环工作事件
        /// </summary>
        event WorkingDelegate OnWorking;

        /// <summary>
        /// 是否已经启动
        /// </summary>
        bool IsStarted { get; }

        /// <summary>
        /// 工作间隔时间
        /// </summary>
        int Intervel { get; set; }

        /// <summary>
        /// 上一次发生的异常
        /// </summary>
        Exception LastException { get; }

        /// <summary>
        /// 启动
        /// </summary>
        void Start();

        /// <summary>
        /// 停止
        /// </summary>
        void Stop();
    }
}
