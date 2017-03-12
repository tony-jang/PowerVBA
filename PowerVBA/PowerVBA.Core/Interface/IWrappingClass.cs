using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Interface
{
    public interface IWrappingClass
    {
        PPTVersion ClassVersion { get; }
    }


    public enum PPTVersion
    {
        /// <summary>
        /// Microsoft PowerPoint 2010
        /// </summary>
        PPT2010,
        /// <summary>
        /// Microsoft PowerPoint 2013
        /// </summary>
        PPT2013,
        /// <summary>
        /// Microsoft PowerPoint 2016
        /// </summary>
        PPT2016
    }
}
