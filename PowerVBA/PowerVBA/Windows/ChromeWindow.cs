using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Interop;
using MINMAXINFO = PowerVBA.Interop.NativeMethods.MINMAXINFO;
using MONITORINFO = PowerVBA.Interop.NativeMethods.MONITORINFO;
using System.Windows;
using System.Windows.Input;
using PowerVBA.Commands;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using WPFExtension;

namespace PowerVBA.Windows
{
    public class ChromeWindow : Window
    {
        

        public static DependencyProperty NoTitleProperty = DependencyHelper.Register();

        public static DependencyProperty IsSubWindowProperty = DependencyHelper.Register(new PropertyMetadata(true));

        public bool IsSubWindow
        {
            get { return (bool)GetValue(IsSubWindowProperty); }
            set { SetValue(IsSubWindowProperty,value); }
        }

        public bool NoTitle
        {
            get { return (bool)GetValue(NoTitleProperty); }
            set { SetValue(NoTitleProperty, value); }
        }

        private bool _IsEnableMove = true;
        public bool IsEnableMove
        {
            get { return _IsEnableMove; }
            set { _IsEnableMove = value; }
        }


        public IntPtr Handle { get; private set; }

        public ChromeWindow()
        {
            this.Style = FindResource("ChromeWindowStyle") as Style;

            CommandBindings.Add(new CommandBinding(WindowSystemCommand.CloseCommand, OnClose, OnCloseExecute));
            CommandBindings.Add(new CommandBinding(WindowSystemCommand.RestoreCommand, OnStateChange, OnStateChangeExecute));
            CommandBindings.Add(new CommandBinding(WindowSystemCommand.MinimizeCommand, OnMiniMize, OnMiniMizeExecute));
        }

        private void OnMiniMizeExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void OnMiniMize(object sender, ExecutedRoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void OnStateChangeExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void OnStateChange(object sender, ExecutedRoutedEventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
            }
            else if (WindowState == WindowState.Normal ||
                     WindowState == WindowState.Minimized)
            {
                WindowState = WindowState.Maximized;
            }
        }

        private void OnCloseExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void OnClose(object sender, ExecutedRoutedEventArgs e)
        {
            this.Close();
        }



        protected override void OnVisualChildrenChanged(DependencyObject visualAdded, DependencyObject visualRemoved)
        {
            base.OnVisualChildrenChanged(visualAdded, visualRemoved);
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            var hwndSource = PresentationSource.FromVisual(this) as HwndSource;

            this.Handle = hwndSource.Handle;

            hwndSource.AddHook(WndProc);
        }

        const int SC_MOVE = 0xF010;

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            switch (msg)
            {
                case NativeMethods.WM_GETMINMAXINFO:
                    WmGetMinMaxInfo(hwnd, ref lParam);
                    break;

                case NativeMethods.WM_SYSCOMMAND:
                    // TODO: Catch maximize, restore
                    if (!IsEnableMove)
                    {
                        int command = wParam.ToInt32() & 0xfff0;
                        if (command == SC_MOVE) handled = true;
                    }

                    break;
            }

            return IntPtr.Zero;
        }



        private void WmGetMinMaxInfo(IntPtr hwnd, ref IntPtr lParam)
        {
            var mmi = Marshal.PtrToStructure<MINMAXINFO>(lParam);
            IntPtr monitor = UnsafeNativeMethods.MonitorFromWindow(hwnd, NativeMethods.MONITOR_DEFAULTTONEAREST);

            if (monitor != IntPtr.Zero)
            {
                var mInfo = new MONITORINFO();
                mInfo.cbSize = Marshal.SizeOf(typeof(MONITORINFO));

                UnsafeNativeMethods.GetMonitorInfo(monitor, ref mInfo);

                mmi.ptMaxPosition.X = Math.Abs(mInfo.rcWork.left - mInfo.rcMonitor.left);
                mmi.ptMaxPosition.Y = Math.Abs(mInfo.rcWork.top - mInfo.rcMonitor.top);
                mmi.ptMaxSize.X = Math.Abs(mInfo.rcWork.right - mInfo.rcWork.left);
                mmi.ptMaxSize.Y = Math.Abs(mInfo.rcWork.bottom - mInfo.rcWork.top);
            }

            Marshal.StructureToPtr(mmi, lParam, true);
        }
    }
}
