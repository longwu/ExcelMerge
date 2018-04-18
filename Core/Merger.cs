using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Configuration;
using Log;
using Microsoft.Office.Interop.Excel;

namespace Core
{
    public class Merger
    {
        private static Merger instance = null;
        private int finished = 0;
        private int total = 0;
        private int currentRowCount = 0;
        private readonly ILoopEngine engine = new LoopEngine();
        private bool isStarted = false;

        /// <summary>
        /// 出现异常
        /// </summary>
        public event Action<string> Errored;
        /// <summary>
        /// 处理进度
        /// </summary>
        public event Action<int, int> Processing;
        /// <summary>
        /// 合并完成
        /// </summary>
        public event Action Completed;

        /// <summary>
        /// 是否合并已经开始
        /// </summary>
        public bool IsStarted
        {
            get { return this.isStarted; }
        }

        /// <summary>
        /// excel缓存路径
        /// </summary>
        private static string cacheDirPath
        {
            get { return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Const.AppName, ".cache"); }
        }

        public static Merger Instance
        {
            get
            {
                if (instance == null)
                    instance = new Merger();
                return instance;
            }
        }

        private Merger()
        {
            engine.Intervel = 500;
            engine.OnWorking += Engine_OnWorking;
            engine.Start();
        }

        private void Engine_OnWorking()
        {
            try
            {
                if (!isStarted)
                    return;

                OnProcessing(this.finished, this.total);
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(ELogLevel.Error, "Merger.Engine_OnWorking", ex.ToString());
            }
        }

        /// <summary>
        /// 启动合并
        /// </summary>
        /// <param name="dirPath"></param>
        public void Start(string dirPath)
        {
            if (this.isStarted)
                return;
            this.isStarted = true;

            this.total = ExcelHelper.GetAvailableExcelCount(dirPath);

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            new AsyncTask<_Workbook>(() =>
            {
                Workbooks workbooks = excel.Workbooks;
                _Workbook result = workbooks.Add();//创建一个空的excel文件

                DirectoryInfo dir = new DirectoryInfo(dirPath);
                FileSystemInfo[] fsList = dir.GetFileSystemInfos();

                foreach (FileSystemInfo fs in fsList)
                {
                    if (fs is FileInfo)
                    {
                        if (fs.Attributes.HasFlag(FileAttributes.Temporary) || fs.Attributes.HasFlag(FileAttributes.Hidden) || fs.Attributes.HasFlag(FileAttributes.NotContentIndexed))
                            continue;

                        if (fs.Extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                            fs.Extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            _Workbook wb = workbooks.Open(fs.FullName);
                            foreach (_Worksheet sheet in wb.Sheets)
                            {
                                Range range = sheet.UsedRange;
                                range.Copy(((_Worksheet)result.Worksheets[1]).Range[string.Format("A{0}", currentRowCount + 1), Missing.Value]);
                                currentRowCount += range.Rows.Count;
                            }
                            wb.Close();
                            this.finished++;//每处理完一个文件 完成数+1
                        }
                    }
                }

                return result;

            }).Run((result, ex) =>
            {
                string path = string.Empty;
                string name = "合并后的excel" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                if (ex != null)
                {
                    OnErrored(ex.Message);
                }
                else
                {
                    path = ExcelHelper.SaveTo(name);
                }

                try
                {
                    if (result != null)
                    {
                        if (!string.IsNullOrEmpty(path))
                            result.SaveAs(path);
                        else
                        {
                            if (!Directory.Exists(cacheDirPath))
                                Directory.CreateDirectory(cacheDirPath);
                            result.SaveAs(System.IO.Path.Combine(cacheDirPath, name));
                        }
                        result.Close();
                    }
                    excel.Quit();

                    if (!string.IsNullOrEmpty(path))
                        ExcelHelper.OpenFileFolder(path);
                }
                catch (Exception ex2)
                {
                    OnErrored(ex2.Message);
                }
                finally
                {
                    this.currentRowCount = 0; //清理上一次记录行号
                    this.finished = 0;
                    this.isStarted = false;
                    OnCompleted();
                }
            });
        }

        private void OnProcessing(int finished, int total)
        {
            if (Processing != null)
                Processing(finished, total);
        }

        private void OnErrored(string error)
        {
            if (Errored != null)
                Errored(error);
        }

        private void OnCompleted()
        {
            if (Completed != null)
                Completed();
        }
    }
}
