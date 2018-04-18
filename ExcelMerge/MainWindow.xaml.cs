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
using Core;
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

        private SynchronizationContext synchronizationContext = null;

        public MainWindow()
        {
            InitializeComponent();

            this.synchronizationContext = SynchronizationContext.Current;

            this.bdDrop.AllowDrop = true;
            this.bdDrop.Drop += GdDrag_Drop;
            this.btnMerge.Click += btnMerge_Click;

            Merger.Instance.Processing += ShowProcessing;
            Merger.Instance.Errored += ShowError;
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
                        this.tbkDragTips.ToolTip = this.tbkDragTips.Text = string.Format("目录:{0}", dirPath);
                        this.tbkCount.ToolTip = this.tbkCount.Text = string.Format("Excel文件数: {0}", ExcelHelper.GetAvailableExcelCount(dirPath));

                        this.gdProgress.Visibility = Visibility.Visible;
                        this.pbProgress.Value = 0;
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

            if (ExcelHelper.GetAvailableExcelCount(dirPath) == 0)
            {
                WinMessageBox.Show("当前文件夹中没有Excel文件");
                return;
            }

            Merger.Instance.Start(dirPath);//合并开始
        }

        /// <summary>
        /// 显示合并进度
        /// </summary>
        /// <param name="finished"></param>
        /// <param name="total"></param>
        private void ShowProcessing(int finished, int total)
        {
            this.synchronizationContext.Send(obj =>
            {
                this.pbProgress.Value = finished;
                this.pbProgress.Maximum = total;

                this.tbkProgress.Text = string.Format("{0}%", Math.Round((double)finished / total, 2) * 100);
            }, null);
        }

        /// <summary>
        /// 显示错误一旦发生
        /// </summary>
        /// <param name="error"></param>
        private void ShowError(string error)
        {
            WinMessageBox.Show(error);
        }
    }
}
