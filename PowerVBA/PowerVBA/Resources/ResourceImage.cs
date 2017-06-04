using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace PowerVBA.Resources
{
    public static class ResourceImage
    {
        private static string BaseURL = "/PowerVBA;Component/Resources/Icon/";

        /// <summary>
        /// 아이콘 이미지를 가져옵니다.
        /// </summary>
        /// <param name="filename">가져올 파일이름입니다. 확장자는 생략합니다.</param>
        /// <returns></returns>
        public static BitmapImage GetIconImage(string filename)
        {
            return new BitmapImage(new Uri(BaseURL + filename + ".png", UriKind.Relative));
        }

        private static string ProgramURL = "/PowerVBA;Component/Resources/ProgramIcon/";

        /// <summary>
        /// 프로그램 아이콘 이미지를 가져옵니다.
        /// </summary>
        /// <param name="filename">가져올 파일이름입니다. 확장자는 생략합니다.</param>
        /// <returns></returns>
        public static BitmapImage GetProgramIconImage(string filename)
        {
            return new BitmapImage(new Uri(ProgramURL + filename + ".ico", UriKind.Relative));
        }

    }
}
