using PowerVBA.Codes;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Global.RegexExpressions;
using PowerVBA.Resources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    /// AddProcedure.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class AddProcedureWindow : ChromeWindow
    {
        CodeEditor Editor;
        LineInfo LineInfo;
        public AddProcedureWindow(CodeEditor editor, string FileName, CodeInfo codeInfo, AddProcedureType AddType)
        {
            InitializeComponent();
            LineInfo = codeInfo.Lines.Where(i=> i.FileName == FileName).FirstOrDefault();

            if (LineInfo == null)
            {
                MessageBox.Show("예외가 발생했습니다!");
                return;
            }

            Editor = editor;


            RoutedCommand AddItem = new RoutedCommand();
            AddItem.InputGestures.Add(new KeyGesture(Key.Escape));

            CommandBinding cb1 = new CommandBinding(AddItem, Comm_Close);

            this.CommandBindings.Add(cb1);

            if (AddType == AddProcedureType.Function) btnFunction.IsChecked = true;
            else if (AddType == AddProcedureType.Property) btnProperty.IsChecked = true;
            else if (AddType == AddProcedureType.Sub) btnSub.IsChecked = true;


            TBName.Focus();
        }

        private void Comm_Close(object sender, ExecutedRoutedEventArgs e)
        {
            Handled = false;
            this.Close();
        }

        public bool Handled;
        public static string[] Procedures = { @"Public Sub %name()

End Sub",
@"Public Function %name()
    
End Function",
@"Public Property Get %name() As Variant
    
End Property

Public Property Let %name(ByVal vNewValue As Variant)
    
End Property",    
        };
        private void AddBtn_Click(object sender, RoutedEventArgs e)
        {
            if (Regex.IsMatch(TBName.Text, CodePattern.Name))
            {
                AddProcedureType t = (btnFunction.IsChecked.Value ? AddProcedureType.Function : (btnProperty.IsChecked.Value ? AddProcedureType.Property : AddProcedureType.Sub));

                Editor.Document.Insert(Editor.Document.GetOffset(LineInfo.LastGlobalVarInt + 1, 0), Procedures[(int)t].Replace("%name", TBName.Text) + "\r\n");
                Editor.ScrollToLine(LineInfo.LastGlobalVarInt + 1);
                Handled = true;
                this.Close();
            }
            else
            {
                MessageBox.Show("명명 규칙에 어긋납니다.");
            }

            
        }

        public new bool ShowDialog()
        {
            base.ShowDialog();
            return Handled;
        }


        public enum AddProcedureType
        {
            Sub,
            Function,
            Property,
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Handled = false;
            this.Close();
        }

        private void TBName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddBtn_Click(this, null);
            }
        }
    }
}
