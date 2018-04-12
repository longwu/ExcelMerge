using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;

namespace ExcelMerge
{
    public class LoopEngine : ILoopEngine
    {
        /// <summary>
        /// 工作事件,用于注册引擎需要执行的方法
        /// </summary>
        public event WorkingDelegate OnWorking;

        private bool isStarted;
        private Exception lastException;
        private bool stopping;

        /// <summary>
        /// 是否已经启动
        /// </summary>
        public bool IsStarted
        {
            get { return this.isStarted; }
        }

        /// <summary>
        /// 引擎执行时间间隔(毫秒ms)
        /// </summary>
        public int Intervel { get; set; }

        /// <summary>
        /// 上一次异常
        /// </summary>
        public Exception LastException
        {
            get { return this.lastException; }
        }

        /// <summary>
        /// 启动引擎
        /// </summary>
        public void Start()
        {
            if (this.isStarted)
                return;

            new Thread(Work) { IsBackground = true }.Start();
            this.isStarted = true;
        }

        /// <summary>
        /// 停止引擎
        /// </summary>
        public void Stop()
        {
            this.stopping = true;
            this.isStarted = false;
        }

        /// <summary>
        /// 循环执行任务
        /// </summary>
        private void Work()
        {
            while (true)
            {
                try
                {
                    if (this.stopping)
                        break;

                    this.OnWorking();
                    Thread.Sleep(this.Intervel);
                }
                catch (Exception ex)
                {
                    this.lastException = ex;
                }
            }
            this.stopping = false;
        }
    }
}
