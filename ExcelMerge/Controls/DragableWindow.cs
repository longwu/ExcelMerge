using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace ExcelMerge
{
    public class DragableWindow : Window
    {
        #region 变量
        private bool _IsClosed = false;
        #endregion

        #region 属性
        public bool IsClosed
        {
            get { return this._IsClosed; }
        }

        new public bool DialogResult
        {
            get { return this.DialogResult; }
            set
            {
                if (!this._IsClosed)
                    base.DialogResult = value;
            }
        }
        #endregion

        public DragableWindow()
        {
            this.Loaded += DragableWindow_Loaded;
            this.MouseLeftButtonDown += DragableWindow_MouseLeftButtonDown;
            this.Closed += WindowBaseDragMove_Closed;
        }

        void DragableWindow_Loaded(object sender, RoutedEventArgs e)
        {

        }

        void DragableWindow_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                this.DragMove();
        }

        void WindowBaseDragMove_Closed(object sender, EventArgs e)
        {
            this._IsClosed = true;
        }

        new void Close()
        {
            if (!this._IsClosed)
                base.Close();
        }

        #region 弹出模态窗口,居屏幕或者父窗体中间显示
        /// <summary>
        /// 弹出模态窗口,居屏幕或者父窗体中间显示
        /// </summary>
        /// <param name="owner">父窗体</param>
        public bool? ShowDialogCenter(Window owner = null)
        {
            if (owner == null)
            {
                this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
            else
            {
                this.Owner = owner;
                this.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            return base.ShowDialog();
        }
        #endregion
    }
}
