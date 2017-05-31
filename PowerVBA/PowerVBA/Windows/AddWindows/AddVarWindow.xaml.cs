using PowerVBA.Codes;
using PowerVBA.Core.AvalonEdit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PowerVBA.Windows.AddWindows
{
    /// <summary>
    /// AddEnumWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class AddVarWindow : ChromeWindow
    {

        CodeEditor Editor;
        LineInfo LineInfo;
        public AddVarWindow(CodeEditor editor, string FileName, CodeInfo codeInfo, bool AddVar)
        {
            InitializeComponent();
            LineInfo = codeInfo.Lines.Where(i => i.FileName == FileName).FirstOrDefault();

            if (LineInfo == null)
            {
                MessageBox.Show("예외가 발생했습니다!");
                return;
            }

            Editor = editor;

            if (AddVar) btnVar.IsChecked = true;
            else btnConst.IsChecked = true;

            RoutedCommand AddItem = new RoutedCommand();
            AddItem.InputGestures.Add(new KeyGesture(Key.Escape));

            CommandBinding cb1 = new CommandBinding(AddItem, Comm_Close);

            this.CommandBindings.Add(cb1);

            TBName.Focus();
        }

        private void Comm_Close(object sender, ExecutedRoutedEventArgs e)
        {
            this.Close();
        }

        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            
            if (btnVar.IsChecked.Value)
            {
                string Accessor = btnDim.IsChecked.Value ? "Dim" : (btnPublic.IsChecked.Value ? "Public" : "Private");
                Editor.Document.Insert(Editor.Document.GetOffset(LineInfo.LastGlobalVarInt + 1, 0), $"{Accessor} {TBName.Text} As Variant\r\n");
            }
            else
            {
                string Accessor = btnDim.IsChecked.Value ? "Const" : "Private Const";
                Editor.Document.Insert(Editor.Document.GetOffset(LineInfo.LastGlobalVarInt + 1, 0), $"{Accessor} {TBName.Text} As Variant = \"Initalize Value\"\r\n");
            }
            this.Close();
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TBName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddBtn_Click(this, null);
            }
        }

        private void btnConst_Checked(object sender, RoutedEventArgs e)
        {
            btnPublic.IsEnabled = !btnConst.IsChecked.Value;
            if (btnConst.IsChecked.Value && btnPublic.IsChecked.Value) btnDim.IsChecked = true;
        }
    }
}
