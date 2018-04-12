using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace ExcelMerge
{
    public class WindowCommands
    {
        public static readonly ICommand CloseCommand = new RelayCommand(c => ((Window)c).Close());
        public static readonly ICommand MiniCommand = new RelayCommand(o => ((Window)o).WindowState = WindowState.Minimized);
        public static readonly ICommand MaxCommand = new RelayCommand(o =>
        {
            if (((Window)o).WindowState == WindowState.Maximized)
                ((Window)o).WindowState = System.Windows.WindowState.Normal;
            else
                ((Window)o).WindowState = System.Windows.WindowState.Maximized;
        });
    }
}
