using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Threading;
using Configuration;
using Log;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelMerge
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : DragableWindow
    {
        private string dirPath = string.Empty;

        private int finished = 0;
        private int total = 0;

        private int currentRowCount = 0;

        private bool isStarted = false;

        private readonly ILoopEngine engine = new LoopEngine();

        private SynchronizationContext synchronizationContext = null;

        private static string cacheDirPath
        {
            get { return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Const.AppName, ".cache"); }
        }

        public MainWindow()
        {
            InitializeComponent();

            this.synchronizationContext = SynchronizationContext.Current;

            this.bdDrop.AllowDrop = true;
            this.bdDrop.Drop += GdDrag_Drop;
            this.btnMerge.Click += btnMerge_Click;
            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
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

                this.synchronizationContext.Send(obj =>
                {
                    this.pbProgress.Value = this.finished;
                    this.pbProgress.Maximum = this.total;

                    this.tbkProgress.Text = string.Format("{0}%", Math.Round((double)finished / total, 2) * 100);
                }, null);
            }
            catch (Exception ex)
            {
                Logger.Instance.Log(ELogLevel.Error, "MainWindow.Engine_OnWorking", ex.ToString());
            }
        }

        private void GdDrag_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    if (Directory.Exists(files[0]))//只允许文件夹
                    {
                        dirPath = files[0];
                        this.total = GetAvailableExcelCount(dirPath);

                        this.tbkDragTips.ToolTip = this.tbkDragTips.Text = string.Format("目录:{0}", dirPath);
                        this.tbkCount.ToolTip = this.tbkCount.Text = string.Format("Excel文件数: {0}", total);

                        this.gdProgress.Visibility = Visibility.Visible;
                        this.pbProgress.Value = this.finished = 0;
                        this.tbkProgress.Text = string.Empty;
                    }
                }
            }
        }

        private void btnMerge_Click(object sender, RoutedEventArgs e)
        {

            if (!Directory.Exists(dirPath))
            {
                WinMessageBox.Show("请选择一个有效的文件夹");
                return;
            }

            if (total == 0)
            {
                WinMessageBox.Show("当前文件夹中没有Excel文件");
                return;
            }

            if (this.isStarted)
                return;

            this.isStarted = true;

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
                        Logger.Instance.Log(ELogLevel.Error, "MainWindow.btnMerge_Click", ex.ToString());
                        WinMessageBox.Show(ex.Message);
                    }
                    else
                    {
                        path = SaveTo(name);
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
                            OpenFileFolder(path);
                    }
                    catch (Exception ex2)
                    {
                        Logger.Instance.Log(ELogLevel.Error, "MainWindow.btnMerge_Click", ex2.ToString());
                        WinMessageBox.Show(ex2.Message);
                    }
                    finally
                    {
                        this.currentRowCount = 0; //清理上一次记录行号
                        this.finished = 0;
                        this.isStarted = false;
                    }
                });
        }

        /// <summary>
        /// 打开文件所在的位置
        /// </summary>
        /// <param name="path"></param>
        private void OpenFileFolder(string path)
        {
            if (File.Exists(path))
            {
                try
                {
                    string args = string.Format("/Select, {0}", path);

                    ProcessStartInfo pfi = new ProcessStartInfo("Explorer.exe", args);
                    pfi.UseShellExecute = false;
                    Process.Start(pfi);
                }
                catch (Exception ex)
                {
                    Logger.Instance.Log(ELogLevel.Error, "MainWindow.OpenFileFolder", ex.ToString());
                    WinMessageBox.Show(ex.Message);
                }
            }
        }

        /// <summary>
        /// 另存为
        /// </summary>
        /// <returns></returns>
        private string SaveTo(string name)
        {
            string path = string.Empty;

            Microsoft.Win32.SaveFileDialog sfd = new Microsoft.Win32.SaveFileDialog();
            sfd.FileName = name;
            sfd.Filter = "Excel 工作簿（*.xlsx）|*.xlsx";
            if (sfd.ShowDialog() == true)
            {
                path = sfd.FileName;
            }

            return path;
        }

        /// <summary>
        /// 获取有效excel文件数量
        /// </summary>
        /// <param name="dirPath"></param>
        /// <returns></returns>
        private int GetAvailableExcelCount(string dirPath)
        {
            int result = 0;
            DirectoryInfo dir = new DirectoryInfo(dirPath);
            FileSystemInfo[] fsList = dir.GetFileSystemInfos();

            foreach (FileSystemInfo fs in fsList)
            {
                if (fs is FileInfo)
                {
                    if (fs.Attributes.HasFlag(FileAttributes.Temporary) ||
                        fs.Attributes.HasFlag(FileAttributes.Hidden) ||
                        fs.Attributes.HasFlag(FileAttributes.NotContentIndexed))
                        continue;

                    if (fs.Extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                        fs.Extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        result++;
                    }
                }
            }
            return result;
        }
    }
}
