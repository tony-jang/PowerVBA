using PowerVBA.Codes;
using PowerVBA.Codes.Parsing;
using PowerVBA.Core.Wrap.WrapBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace PowerVBA.Windows
{
    /// <summary>
    /// ErrorWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ErrorWindow : ChromeWindow
    {
        public ErrorWindow(List<VBComponentWrappingBase> Components)
        {
            InitializeComponent();

            try
            {
                var info = new CodeInfo();

                VBAParser p = new VBAParser(info);


                Components.ForEach(i => info.AddFile(i.CompName));

                p.Parse(Components.Select(i => (i.CompName, i.Code)).ToList());
                lvUsers.ItemsSource = info.ErrorList;

                RunFileCount.Text = Components.Count.ToString();

                RunAllErrCount.Text = info.ErrorList.Count.ToString();
                RunErrorCount.Text = info.ErrorList.Where(i => i.ErrorType == ErrorType.Error).Count().ToString();
                RunWarningCount.Text = info.ErrorList.Where(i => i.ErrorType == ErrorType.Warning).Count().ToString();

                CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(lvUsers.ItemsSource);
                PropertyGroupDescription groupDescription = new PropertyGroupDescription("FileName");
                view.GroupDescriptions.Add(groupDescription);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            
        }
        
    }
}
