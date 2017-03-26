using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace PowerVBA.Codes.Attributes
{
    abstract class BaseErrorAttribute : Attribute
    {
        public CultureInfo MessageCulture { get; set; }
        public virtual string ErrorMessage { get; }
        public BaseErrorAttribute(string ErrorMessage)
        {
            this.ErrorMessage = ErrorMessage;
        }
    }
}
