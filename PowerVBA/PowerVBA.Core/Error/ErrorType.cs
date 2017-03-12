using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Error
{
    public enum ErrorType
    {
        /// <summary>
        /// 오류
        /// </summary>
        Error,
        /// <summary>
        /// 경고 메세지
        /// </summary>
        Warning,
        /// <summary>
        /// 안내 메세지
        /// </summary>
        Message
    }
}
