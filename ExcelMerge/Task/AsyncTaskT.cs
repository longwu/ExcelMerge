using System;
using System.Threading;

namespace ExcelMerge
{
    /// <summary>
    /// 声明带泛型返回值的委托
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    public delegate T Func<T>();

    /// <summary>
    /// 声明带参数无返回值的委托
    /// </summary>
    /// <typeparam name="T1"></typeparam>
    /// <typeparam name="T2"></typeparam>
    /// <param name="arg1"></param>
    /// <param name="arg2"></param>
    public delegate void Action<T1, T2>(T1 arg1, T2 arg2);

    public class AsyncTask<T>
    {
        /// <summary>
        /// 异步方法,带返回参数
        /// </summary>
        private event Func<T> Func = null;

        /// <summary>
        ///  异步任务执行过程中的回调方法
        /// </summary>
        private event Action<T, Exception> ActionEx = null;

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
        /// <param name="func">需要异步执行的方法,带返回值</param>
        public AsyncTask(Func<T> func)
        {
            syncContext = SynchronizationContext.Current;
            this.Func += func;
        }

        /// <summary>
        /// 异步执行
        /// </summary>
        /// <param name="actionEx">任务完成后的同步回调方法</param>
        public void Run(Action<T, Exception> actionEx)
        {
            this.ActionEx += actionEx;
            this.Run();
        }

        /// <summary>
        /// 开始异步执行方法
        /// </summary>
        private void Run()
        {
            if (this.Func != null)
            {
                ThreadPool.QueueUserWorkItem(obj =>
                {
                    T result = default(T);
                    try
                    {
                        this.isStarted = true;
                        result = this.Func.Invoke(); //执行异步方法
                    }
                    catch (Exception ex)
                    {
                        this.exception = ex;
                        this.isFaulted = true; //执行任务失败
                    }
                    finally
                    {
                        if (!this.isCancelled)
                        {
                            this.isEnded = true;
                            ContinueWith(result,this.exception); //异步任务完成后进行同步回调
                        }
                    }
                });
            }
        }

        /// <summary>
        /// 异步任务完成后继续同步回调
        /// </summary>
        /// <param name="result">任务返回值</param>
        /// <param name="ex">任务异常</param>
        private void ContinueWith(T result, Exception ex)
        {
            if (this.ActionEx != null)
            {
                this.syncContext.Send(obj =>
                {
                    if (this.ActionEx != null)
                    {
                        this.ActionEx.Invoke(result, ex);//同步回调
                    }
                }, null);
            }
        }
    }
}
