using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Connector
{
    public enum PPTVersion
    {
        /// <summary>
        /// Microsoft PowerPoint 2016 버전입니다. (16.0)
        /// </summary>
        [Description("2016")]
        PPT2016 = 16,
        /// <summary>
        /// Microsoft PowerPoint 2013 버전입니다. (15.0)
        /// </summary>
        [Description("2013")]
        PPT2013 = 15,
        /// <summary>
        /// Microsoft PowerPoint 2010 버전입니다. (14.0)
        /// </summary>
        [Description("2010")]
        PPT2010 = 14,
        /// <summary>
        /// Microsoft PowerPoint 2007 버전입니다. (12.0)
        /// </summary>
        [Description("2007")]
        PPT2007 = 12,
        /// <summary>
        /// Microsoft PowerPoint 2003 버전입니다. (11.0)
        /// </summary>
        [Description("2003")]
        PPT2003 = 11,
        /// <summary>
        /// Microsoft PowerPoint XP 버전입니다. (10.0)
        /// </summary>
        [Description("XP")]
        PPTXP = 10,
        /// <summary>
        /// Microsoft PowerPoint 2000 버전입니다. (9.0)
        /// </summary>
        [Description("2000")]
        PPT2000 = 9,
        /// <summary>
        /// Microsoft PowerPoint 98 버전입니다. (8.0)
        /// </summary>
        [Description("98")]
        PPT98 = 8,
        /// <summary>
        /// Microsoft PowerPoint 97 버전입니다. (7.0)
        /// </summary>
        [Description("97")]
        PPT97 = 7,

        /// <summary>
        /// 알 수 없는 버전입니다.
        /// </summary>
        Unknown = 0
    }
}