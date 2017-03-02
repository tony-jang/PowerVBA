using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace PowerVBA.Resources
{
    static class ResourceImage
    {
        private static string BaseURL = "/PowerVBA;Component/Resources/Icon/";
        public static BitmapImage GetIconImage(string filename)
        {
            return new BitmapImage(new Uri(BaseURL + filename + ".png", UriKind.Relative));
        }

        private static string ProgramURL = "/PowerVBA;Component/Resources/ProgramIcon/";
        public static BitmapImage GetProgramIconImage(string filename)
        {
            return new BitmapImage(new Uri(ProgramURL + filename + ".ico", UriKind.Relative));
        }

    }
}
