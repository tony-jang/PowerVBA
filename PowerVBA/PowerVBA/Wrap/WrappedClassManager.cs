using PowerVBA.Core.Connector;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.V2010.Connector;
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

        public static V2013.WrapClass.ApplicationWrapping ToApplication2013(this ApplicationWrappingBase appbase)
        {
            return (V2013.WrapClass.ApplicationWrapping)appbase;
        }

        public static V2013.WrapClass.PresentationWrapping ToPresentation2013(this PresentationWrappingBase pptbase)
        {
            return (V2013.WrapClass.PresentationWrapping)pptbase;
        }

        public static V2013.WrapClass.VBComponentWrapping ToVBComponent2013(this VBComponentWrappingBase vbcompbase)
        {
            return (V2013.WrapClass.VBComponentWrapping)vbcompbase;
        }

        public static V2013.WrapClass.VBProjectWrapping ToVBProj2013(this VBProjectWrappingBase vbprojbase)
        {
            return (V2013.WrapClass.VBProjectWrapping)vbprojbase;
        }

        public static V2013.WrapClass.SlideWrapping ToSlide2013(this SlideWrappingBase slidebase)
        {
            return (V2013.WrapClass.SlideWrapping)slidebase;
        }

        public static V2013.WrapClass.ShapeWrapping ToShape2013(this ShapeWrappingBase shapebase)
        {
            return (V2013.WrapClass.ShapeWrapping)shapebase;
        }

        public static PPTConnector2013 ToConnector2013(this PPTConnectorBase connbase)
        {
            return ((PPTConnector2013)connbase);
        }

        #endregion


        #region [  Powerpoint 2010  ]

        public static V2010.WrapClass.ApplicationWrapping ToApplication2010(this ApplicationWrappingBase appbase)
        {
            return (V2010.WrapClass.ApplicationWrapping)appbase;
        }

        public static V2010.WrapClass.PresentationWrapping ToPresentation2010(this PresentationWrappingBase pptbase)
        {
            return (V2010.WrapClass.PresentationWrapping)pptbase;
        }

        public static V2010.WrapClass.VBComponentWrapping ToVBComponent2010(this VBComponentWrappingBase vbcompbase)
        {
            return (V2010.WrapClass.VBComponentWrapping)vbcompbase;
        }

        public static V2010.WrapClass.VBProjectWrapping ToVBProj2010(this VBProjectWrappingBase vbprojbase)
        {
            return (V2010.WrapClass.VBProjectWrapping)vbprojbase;
        }

        public static V2010.WrapClass.SlideWrapping ToSlide2010(this SlideWrappingBase slidebase)
        {
            return (V2010.WrapClass.SlideWrapping)slidebase;
        }

        public static V2010.WrapClass.ShapeWrapping ToShape2010(this ShapeWrappingBase shapebase)
        {
            return (V2010.WrapClass.ShapeWrapping)shapebase;
        }

        public static PPTConnector2010 ToConnector2010(this PPTConnectorBase connbase)
        {
            return ((PPTConnector2010)connbase);
        }

        #endregion
    }
}
