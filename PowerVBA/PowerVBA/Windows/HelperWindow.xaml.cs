using PowerVBA.Controls.Customize;
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

namespace PowerVBA.Windows
{
    /// <summary>
    /// HelperWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class HelperWindow : ChromeWindow
    {
        List<Grid> cache = new List<Grid>();
        public HelperWindow()
        {
            InitializeComponent();

            treeList.SelectedItemChanged += TreeList_SelectedItemChanged;
            
            // Add Help Tree

            // 기초가 되는 첫 아이템
            AddTree(treeList, "첫 도움말", "FirstHelp");
            AddTree(treeList, "PowerVBA 도움말", "BasicHelp");
            AddTree(treeList, "컴포넌트 (파일) 추가/제거", "ComponentHelp");
            AddTree(treeList, "PowerPoint, PowerVBA 차이점", "PPTDiffHelp");
            AddTree(treeList, "참조 추가하기", "ReferenceHelp");
            AddTree(treeList, "함수 선언하기", "FunctionHelp");
            AddTree(treeList, "코드 에디터 기능 목록", "CodeEditorHelp");
            AddTree(treeList, "미리 정의된 함수 추가/제거", "PreDecFuncHelp");
            AddTree(treeList, "트리거 추가/제거", "TriggerHelp");
            AddTree(treeList, "VBA 형식", "TypeHelp");
            AddTree(treeList, "코드 분석", "CodeAnalysisHelp");
            AddTree(treeList, "명명 규칙", "NamingHelp");
            AddTree(treeList, "도형 탐색기", "ShapeExplorerHelp");


            MoveHelpContext("FirstHelp");
        }

        private void TreeList_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            MoveHelpContext(((TreeViewItem)treeList.SelectedItem).Tag.ToString());
        }

        private void MoveHelpContext(object sender)
        {
            ImageButton btn = sender as ImageButton;

            string DocumentName = (string)btn.Tag;

            MoveHelpContext(DocumentName);
        }

        private TreeViewItem AddTree(ItemsControl control, string ViewName, string MoveLink)
        {
            var itm = new TreeViewItem()
            {
                Header = ViewName,
                Tag = MoveLink
            };
            if (control == null)
            {
                return null;
            }
            control.Items.Add(itm);

            return itm;
        }
        
        private void I_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink hl = (Hyperlink)sender;

            MoveHelpContext(hl.TargetName);
        }


        Stack<string> UndoStack = new Stack<string>();
        Stack<string> RedoStack = new Stack<string>();
        public void MoveNext()
        {
            if (RedoStack.Count == 0) return;
            MoveHelpContext(RedoStack.Pop(), MoveCallType.Redo);
        }
        public void MovePrev()
        {
            if (UndoStack.Count == 0) return;
            string value = UndoStack.Pop();
            //RedoStack.Push(value);
            MoveHelpContext(value, MoveCallType.Undo);
        }

        private void MoveHelpContext(string DocumentName, MoveCallType OriginalCall)
        {
            if (OriginalCall == MoveCallType.General)
            {
                RedoStack.Clear();
            }
            if (HelpFrame.Tag != null)
            {
                var value = HelpFrame.Tag.ToString();
                switch (OriginalCall)
                {
                    case MoveCallType.General:
                    case MoveCallType.Redo:
                        if (UndoStack.Count == 0 || UndoStack.Peek() != value) UndoStack.Push(value);
                        break;
                    case MoveCallType.Undo:
                        if (RedoStack.Count == 0 || RedoStack.Peek() != value) RedoStack.Push(HelpFrame.Tag.ToString());
                        break;
                }
            }

            try
            {
                Grid newDoc = FindResource(DocumentName) as Grid;
                if (!cache.Contains(newDoc))
                {
                    cache.Add(newDoc);
                    List<object> obj = GetAllChildrens(newDoc);
                    List<ImageButton> buttons = obj.Where(i => i.GetType() == typeof(ImageButton))
                                                                       .Cast<ImageButton>().ToList();

                    List<Hyperlink> hyperlinks = obj.Where(i => i.GetType() == typeof(Hyperlink))
                                                                       .Cast<Hyperlink>().ToList();

                    buttons.ForEach(i => i.ButtonClick += MoveHelpContext);
                    hyperlinks.ForEach(i => i.Click += I_Click);

                }

                HelpFrame.Content = newDoc;
                HelpFrame.Tag = DocumentName;
                runViewDoc.Text = DocumentName;
            }
            catch (ResourceReferenceKeyNotFoundException)
            {
                MessageBox.Show("도움말 링크가 잘못 되었습니다. 오류 문의에 문의해주세요.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("알 수 없는 오류가 발생했습니다." + Environment.NewLine + Environment.NewLine + ex.ToString());
            }
        }
        public void MoveHelpContext(string DocumentName)
        {
            MoveHelpContext(DocumentName, MoveCallType.General);
        }

        public List<object> GetAllChildrens(Panel panel)
        {
            List<object> Child = new List<object>();
            foreach(UIElement c in panel.Children)
            {
                Child.Add(c);
                if (c.GetType() == (typeof(Panel)) || c.GetType().IsSubclassOf(typeof(Panel)))
                {
                    Child.AddRange(GetAllChildrens(c as Panel));
                }
                else if (c.GetType() == typeof(TextBlock) || c.GetType().IsSubclassOf(typeof(TextBlock)))
                {
                    Child.AddRange(((TextBlock)c).Inlines.ToList());
                }
            }
            return Child;
        }

        private void btnHome_Click(object sender, RoutedEventArgs e)
        {
            MoveHelpContext("FirstHelp");
        }

        private void btnBack_Click(object sender, RoutedEventArgs e)
        {
            MovePrev();
        }

        private void btnFront_Click(object sender, RoutedEventArgs e)
        {
            MoveNext();
        }
    }


    public enum MoveCallType
    {
        General,
        Undo,
        Redo,
    }
}
