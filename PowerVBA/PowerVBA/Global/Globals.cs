using ICSharpCode.AvalonEdit.Highlighting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Globalization;
using System.Windows.Controls;
using ICSharpCode.AvalonEdit;

namespace PowerVBA.Global
{
    static class Globals
    {
        
        public static IHighlightingDefinition highlightingDefintion;

        
        public static Type[] GetTypesInNamespace(Assembly assembly, string nameSpace)
        {
            return assembly.GetTypes().Where(t => string.Equals(t.Namespace, nameSpace, StringComparison.Ordinal)).ToArray();
        }


        public static Size MeasureString(string candidate, TextEditor editor)
        {
            var formattedText = new FormattedText(
                candidate,
                CultureInfo.CurrentUICulture,
                FlowDirection.LeftToRight,
                new Typeface(editor.FontFamily, editor.FontStyle, editor.FontWeight, editor.FontStretch),
                editor.FontSize,
                Brushes.Black);

            var size =  new Size(formattedText.Width, formattedText.Height);

            if (size.Width == 0) size.Width = 5;

            return size;
        }
    }
}
