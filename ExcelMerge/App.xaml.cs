using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using Configuration;
using Log;
using Microsoft.Shell;
using Microsoft.Win32;

namespace ExcelMerge
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application, ISingleInstanceApp
    {
        public App()
        {
            this.Startup += App_Startup;

            this.Exit += App_Exit;
        }

        void App_Startup(object sender, StartupEventArgs e)
        {
            //初始化样式
            InitStyle();

            //写入运行路径
            SetAppPath();

            //创建应用程序缓存目录
            CreateAppDir();

            //初始化日志
            InitialLog();

            MainWindow main = new MainWindow();
            main.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            main.Show();
        }

        private void App_Exit(object sender, ExitEventArgs e)
        {
            try
            {
                Logger.Instance.Dispose(); //一定要释放,否则程序无法退出
            }
            catch (Exception exception)
            {

            }
        }

        /// <summary>
        /// 初始化日志
        /// </summary>
        private void InitialLog()
        {
            Logger.Instance.LogLevel = ELogLevel.Error;
        }

        /// <summary>
        /// 创建应用程序缓存目录
        /// </summary>
        private void CreateAppDir()
        {
            string rootDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Const.AppName);
            if (!Directory.Exists(rootDir))
            {
                Directory.CreateDirectory(rootDir);
            }
        }

        #region 样式
        /// <summary>
        /// 初始化样式
        /// </summary>
        private void InitStyle()
        {
            base.Resources.MergedDictionaries.Add(Application.LoadComponent(new Uri("/Style/Geometry.xaml", UriKind.Relative)) as ResourceDictionary);
            base.Resources.MergedDictionaries.Add(Application.LoadComponent(new Uri("/Style/Brush.xaml", UriKind.Relative)) as ResourceDictionary);
            base.Resources.MergedDictionaries.Add(Application.LoadComponent(new Uri("/Style/Theme.xaml", UriKind.Relative)) as ResourceDictionary);
        }
        #endregion

        #region 设置运行目录
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
                Logger.Instance.Log(ELogLevel.Error, "App.SetAppPath", "写入程序运行路径失败");
                Logger.Instance.Log(ELogLevel.Error, "App.SetAppPath", ex.ToString());
            }
        }

        #endregion

        #region 单例
        /// <summary>
        /// 程序运行入口
        /// </summary>
        /// <param name="args"></param>
        [STAThread]
        public static void Main(string[] args)
        {
            if (SingleInstance<App>.InitializeAsFirstInstance(Const.AppId.ToString()))
            {
                App application = new App();
                application.InitializeComponent();
                application.Run();

                SingleInstance<App>.Cleanup();
            }
        }

        /// <summary>
        /// 二次启动函数
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        public bool SignalExternalCommandLineArgs(IList<string> args)
        {
            if (this.MainWindow.WindowState == WindowState.Minimized)
            {
                this.MainWindow.WindowState = WindowState.Normal;
                this.MainWindow.ShowInTaskbar = true;
            }

            this.MainWindow.Activate();
            return true;
        }
        #endregion
    }
}
