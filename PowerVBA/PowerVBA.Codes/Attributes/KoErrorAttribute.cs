using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.Attributes
{
    /// <summary>
    /// <see cref="BaseErrorAttribute"/>에서 상속 받은 한국어 오류를 나타냅니다.
    /// </summary>
    class KoErrorAttribute : BaseErrorAttribute
    {
        public KoErrorAttribute(string ErrorMessage) : base(ErrorMessage)
        {
            MessageCulture = new CultureInfo("ko-KR");
        }
    }
}
