using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.V2013.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Wrap
{
    static class WrappedClassManager
    {
        #region [  Powerpoint 2013  ]

        public static V2013.Wrap.WrapClass.ApplicationWrapping ToApplication2013(this ApplicationWrappingBase appbase)
        {
            return (V2013.Wrap.WrapClass.ApplicationWrapping)appbase;
        }

        public static V2013.Wrap.WrapClass.PresentationWrapping ToPresentation2013(this PresentationWrappingBase pptbase)
        {
            return (V2013.Wrap.WrapClass.PresentationWrapping)pptbase;
        }

        public static V2013.Wrap.WrapClass.VBComponentWrapping ToVBComponent2013(this VBComponentWrappingBase vbcompbase)
        {
            return (V2013.Wrap.WrapClass.VBComponentWrapping)vbcompbase;
        }

        public static V2013.Wrap.WrapClass.VBProjectWrapping ToVBProj2013(this VBProjectWrappingBase vbprojbase)
        {
            return (V2013.Wrap.WrapClass.VBProjectWrapping)vbprojbase;
        }

        public static V2013.Wrap.WrapClass.SlideWrapping ToSlide2013(this SlideWrappingBase slidebase)
        {
            return (V2013.Wrap.WrapClass.SlideWrapping)slidebase;
        }

        public static V2013.Wrap.WrapClass.ShapeWrapping ToShape2013(this ShapeWrappingBase shapebase)
        {
            return (V2013.Wrap.WrapClass.ShapeWrapping)shapebase;
        }

        public static PPTConnector2013 ToConnector2013(this PPTConnectorBase connbase)
        {
            return ((PPTConnector2013)connbase);
        }

        #endregion
    }
}
