using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Configuration;
using Microsoft.Win32;

namespace Install
{
    class Program
    {
        static void Main(string[] args)
        {
            //卸载 
            if (args.Length == 1 && args[0].Equals("-uninstall", StringComparison.OrdinalIgnoreCase))
            {
                KillClientProcess();
            }
            else//安装
            {
                KillClientProcess();

                SetAppPath();//将安装路径写入注册表
            }
        }

        /// <summary>
        /// 关掉应用程序进程
        /// </summary>
        private static void KillClientProcess()
        {
            //卸载
            foreach (Process p in Process.GetProcesses())
            {
                if (p.ProcessName.Equals("ExcelMerge", StringComparison.OrdinalIgnoreCase))
                {
                    p.Kill();
                    p.WaitForExit();
                }
            }
        }

        /// <summary>
        /// 将运行文件路径写入注册表
        /// </summary>
        private static void SetAppPath()
        {
            try
            {
                RegistryKey software = Registry.CurrentUser.OpenSubKey("Software", true);
                RegistryKey im = software.OpenSubKey(Const.AppName, true);
                if (im == null)
                {
                    im = software.CreateSubKey(Const.AppName);
                }
                im.SetValue("Path", System.Reflection.Assembly.GetExecutingAssembly().Location);
                im.Close();
                software.Close();
            }
            catch (Exception ex)
            {
            }
        }
    }
}
