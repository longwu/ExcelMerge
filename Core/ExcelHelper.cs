using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Core
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 打开文件所在的位置
        /// </summary>
        /// <param name="path"></param>
        public static void OpenFileFolder(string path)
        {
            if (File.Exists(path))
            {
                string args = string.Format("/Select, {0}", path);

                ProcessStartInfo pfi = new ProcessStartInfo("Explorer.exe", args);
                pfi.UseShellExecute = false;
                Process.Start(pfi);
            }
        }


        /// <summary>
        /// 另存为
        /// </summary>
        /// <returns></returns>
        public static string SaveTo(string name)
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
        /// <param name="dirpath"></param>
        /// <returns></returns>
        public static int GetAvailableExcelCount(string dirpath)
        {
            int result = 0;
            DirectoryInfo dir = new DirectoryInfo(dirpath);
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
