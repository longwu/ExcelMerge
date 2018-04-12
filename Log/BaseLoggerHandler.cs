using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Collections;
using System.Diagnostics;

namespace Log
{
    /// <summary>
    /// 日志处理类
    /// </summary>
    public abstract class BaseLoggerHandler
    {
        private bool _Alive = false;
        private Queue _Queue = null;
        private Thread _Thread = null;
        protected string[] _LogLevelDescriptors = null;

        public BaseLoggerHandler()
        {
            this._Queue = Queue.Synchronized(new Queue(100));
            this.Start();
        }

        /// <summary>
        /// 开始执行日志处理程序
        /// </summary>
        public void Start()
        {
            //是否准备好
            if (this._Alive)
                return;
              
            this._Alive = true;
            this._Thread = new Thread(new ThreadStart(LogMessages));
            this._Thread.Start();
        }

        /// <summary>
        /// 停止
        /// </summary>
        public void Shutdown()
        {
            if (!this._Alive) return;

            this._Alive = false;
            Monitor.Enter(this._Queue);
            Monitor.PulseAll(this._Queue);
            Monitor.Exit(this._Queue);
            while (this._Thread != null && this._Thread.IsAlive)
            {
                Thread.Sleep(10);
            }
            //this.Abort();
        }

        /// <summary>
        /// 关闭线程
        /// </summary>
        public void Abort()
        {
            if (!this._Alive)
                return;

            try
            {
                this._Thread.Abort();
            }
            catch
            { }
        }

        /// <summary>
        /// 记录日志
        /// </summary>
        protected void LogMessages()
        {
            while (this._Alive)
            {
                while ((this._Queue.Count != 0) && this._Alive)
                {
                    this.Log((LoggerMessage)this._Queue.Dequeue());
                }

                if ((this._Alive) && (this._Queue.Count == 0))
                {
                    Monitor.Enter(this._Queue);
                    if (this._Queue.Count == 0)
                        Monitor.Wait(this._Queue);
                    Monitor.Exit(this._Queue);
                }
                Thread.Sleep(10);
            }

            while (this._Queue.Count != 0)
            {
                this.Log((LoggerMessage)this._Queue.Dequeue());
            }

            this.OnShutdown();
            this._Thread = null;
        }

        /// <summary>
        /// 记录日志的方法
        /// </summary>
        /// <param name="tag"></param>
        /// <param name="level"></param>
        /// <param name="levelthis._desc"></param>
        /// <param name="message"></param>
        public void Log(ELogLevel level, string tag, string message)
        {
            if (!this._Alive)
                return;

            LoggerMessage msg = new LoggerMessage();
            msg.Message = message;
            msg.Tag = tag;
            msg.Level = level;
            msg.Time = System.DateTime.Now;

            this._Queue.Enqueue(msg);

            Monitor.Enter(this._Queue);
            Monitor.PulseAll(this._Queue);
            Monitor.Exit(this._Queue);
        }

        /// <summary>
        /// 实现的记录日志文件的方法
        /// </summary>
        /// <param name="message"></param>
        protected abstract void Log(LoggerMessage message);

        /// <summary>
        /// 关闭记录日志文件的方法
        /// </summary>
        protected abstract void OnShutdown();
    }
}
