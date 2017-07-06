using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using PowerVBA.Codes;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Wrap;

namespace PowerVBA.Windows.AddWindows
{
    /// <summary>
    /// AddTriggerWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class AddTriggerWindow : ChromeWindow
    {
        public AddTriggerWindow(bool IsMouseOver, CodeEditor editor, CodeInfo codeInfo, string FileName)
        {
            InitializeComponent();

            this.FileName = FileName;

            Editor = editor;
            CodeInfo = codeInfo;

            if (IsMouseOver)
            { btnMouseOver.IsChecked = true; }
            else
            { btnMouseClick.IsChecked = true; }
        }

        PPTConnectorBase Connector;
        CodeEditor Editor;
        CodeInfo CodeInfo;
        ShapeWrappingBase ReturnShape;

        bool Handled = false;
        string FileName;

        public bool ShowDialog(PPTConnectorBase connector)
        {
            Connector = connector;

            RefreshBtn_Click(this, null);

            base.ShowDialog();

            if (!Handled)
            {
                CBSlide.IsEnabled = true;
                AddBtn.IsEnabled = true;
                CBConnShapes.IsEnabled = true;
                
                return false;
            }
            return true;
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Handled = false;
            this.Close();
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var vInt = CodeInfo.Lines.Where(i => i.FileName == FileName).First().LastGlobalVarInt;

                string MethodName =
                    $@"Slide{CBSlide.SelectedValue.ToString()
                    }_{CBConnShapes.SelectedValue.ToString().Replace(' ', '_')
                    }_{(btnMouseClick.IsChecked.Value ? "MouseClick" : "MouseOver")}";

                Editor.Document.Insert(Editor.Document.GetOffset(vInt + 1, 0),
                    $@"Public Sub {MethodName}()
    
End Sub{Environment.NewLine}");

                Editor.Save();

                Connector.AddMacro(MethodName, ReturnShape, btnMouseOver.IsChecked.Value);

                Handled = true;
            }
            catch (Exception)
            {
                Handled = false;
            }
            
            this.Close();
        }

        private void CBSlide_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CBSlide.SelectedIndex == -1) return;

            CBConnShapes.Items.Clear();

            foreach (var s in Connector.Shapes(int.Parse(CBSlide.SelectedValue.ToString()))) CBConnShapes.Items.Add(s.ToString());

            if (CBConnShapes.Items.Count == 0)
            {
                CBConnShapes.IsEnabled = false;
                AddBtn.IsEnabled = false;
            }
            else
            {
                CBConnShapes.SelectedIndex = 0;
                CBConnShapes.IsEnabled = true;
                AddBtn.IsEnabled = true;
            }
        }

        private void CBConnShapes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CBConnShapes.SelectedIndex == -1) return;

            ReturnShape = Connector.Shapes().Where(i => i.ToString() == CBConnShapes.SelectedValue.ToString()).First();
        }

        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            CBSlide.Items.Clear();
            for (int i = 1; i <= Connector.SlideCount; i++) CBSlide.Items.Add(i);
            CBConnShapes.Items.Clear();

            if (CBSlide.Items.Count == 0)
            {
                AddBtn.IsEnabled = false;
                CBSlide.IsEnabled = false;
                CBConnShapes.IsEnabled = false;
            }

            CBSlide.SelectedIndex = 0;
        }
    }
}
