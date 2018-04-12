using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelMerge
{
    /// <summary>
    /// Interaction logic for WinMessageBox.xaml
    /// </summary>
    public partial class WinMessageBox : DragableWindow
    {
        public WinMessageBox(string content)
        {
            InitializeComponent();

            this.tbkContent.Text = content;
            this.btnClose.Click += btnClose_Click;
        }

        void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 显示对话框
        /// </summary>
        /// <param name="content">内容</param>
        /// <param name="owner">父窗体</param>
        public static void Show(string content, Window owner = null)
        {
            WinMessageBox winMessage = new WinMessageBox(content);

            if (owner != null)
            {
                winMessage.Owner = owner;
                winMessage.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else
            {
                if (Application.Current.MainWindow != null)
                {
                    winMessage.Owner = Application.Current.MainWindow;
                    winMessage.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                }
                else
                {
                    winMessage.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                }
            }
            winMessage.ShowDialog();
        }
    }
}
