using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace PowerVBA
{

    /// <summary>
    /// App.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class App : Application
    {
        //[DllImport("user32.dll")]
        //private static extern bool SetForegroundWindow(IntPtr hWnd);

        //Mutex _mutex = null;
        
        //protected override void OnStartup(StartupEventArgs e)
        //{
        //    _mutex = new Mutex(true, "PowerVBA", out bool isCreatedNew);
        //    if (isCreatedNew)
        //    {
        //        base.OnStartup(e);
        //    }
        //    else
        //    {
        //        SetForegroundWindow(Mutex.OpenExisting("PowerVBA").SafeWaitHandle.DangerousGetHandle());
        //        Application.Current.Shutdown();
        //    }
        //}
    }
}
