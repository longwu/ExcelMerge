using System;
using System.Threading;

namespace ExcelMerge
{
    /// <summary>
    /// 声明无参数委托
    /// </summary>
    public delegate void Action();

    /// <summary>
    /// 带泛型参数的委托
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="args"></param>
    public delegate void Action<T>(T args);

    /// <summary>
    /// 一个异步任务类
    /// </summary>
    public class AsyncTask
    {
        /// <summary>
        /// 声明任务事件
        /// </summary>
        private event Action Action = null;

        /// <summary>
        /// 声明带异常参数的任务事件
        /// </summary>
        private event Action<Exception> ActionEx = null;

        /// <summary>
        /// 任务是否已经启动
        /// </summary>
        private bool isStarted = false;

        /// <summary>
        /// 任务是否已经结束
        /// </summary>
        private bool isEnded = false;

        /// <summary>
        /// 任务是否取消了
        /// </summary>
        private bool isCancelled = false;

        /// <summary>
        /// 任务是否出错了
        /// </summary>
        private bool isFaulted = false;

        /// <summary>
        /// 执行异步中产生的异常
        /// </summary>
        private Exception exception;

        /// <summary>
        /// 线程调度
        /// </summary>
        private SynchronizationContext syncContext = null;

        /// <summary>
        /// 任务是否已经启动
        /// </summary>
        public bool IsStarted
        {
            get { return this.isStarted; }
        }

        /// <summary>
        /// 任务是否已经结束
        /// </summary>
        public bool IsEnded
        {
            get { return this.isEnded; }
        }

        /// <summary>
        /// 任务是否取消了
        /// </summary>
        public bool IsCancelled
        {
            get { return this.isCancelled; }
        }

        /// <summary>
        /// 任务是否出错了
        /// </summary>
        public bool IsFaulted
        {
            get { return this.isFaulted; }
        }

        /// <summary>
        /// 获取任务执行过程中产生的异常
        /// </summary>
        public Exception Exception
        {
            get { return this.exception; }
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="action">需要异步执行的方法</param>
        public AsyncTask(Action action)
        {
            syncContext = SynchronizationContext.Current;
            this.Action += action;
        }

        /// <summary>
        /// 异步执行
        /// </summary>
        /// <param name="actionEx">任务完成后的同步回调方法</param>
        public void Run(Action<Exception> actionEx)
        {
            this.ActionEx += actionEx;
            this.Run();
        }

        /// <summary>
        /// 开始异步执行方法
        /// </summary>
        private void Run()
        {
            if (this.Action != null)
            {
                ThreadPool.QueueUserWorkItem(obj =>
                {
                    try
                    {
                        this.isStarted = true;
                        this.Action.Invoke(); //执行方法
                    }
                    catch (Exception ex)
                    {
                        this.exception = ex;
                        this.isFaulted = true; //执行失败
                    }
                    finally
                    {
                        if (!this.isCancelled)
                        {
                            this.isEnded = true;
                            ContinueWith(this.exception); //异步方法完成后进行同步回调
                        }
                    }
                });
            }
        }

        /// <summary>
        /// 异步任务完成后继续同步回调
        /// </summary>
        /// <param name="ex">异常信息</param>
        private void ContinueWith(Exception ex)
        {
            if (this.ActionEx != null)
            {
                this.syncContext.Send(obj =>
                {
                    if (this.ActionEx != null)
                    {
                        this.ActionEx.Invoke(ex);//同步回调
                    }
                }, null);
            }
        }

        /// <summary>
        /// 取消任务
        /// </summary>
        public void Cancel()
        {
            this.isCancelled = true;
        }
    }
}
