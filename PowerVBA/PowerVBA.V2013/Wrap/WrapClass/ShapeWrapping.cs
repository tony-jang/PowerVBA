using Microsoft.Office.Interop.PowerPoint;
using core=Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerVBA.Core.Interface;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;

namespace PowerVBA.V2013.Wrap.WrapClass
{
    [Wrapped(typeof(Shape))]
    public class ShapeWrapping : ShapeWrappingBase
    {

        public Shape Shape { get; }

        public ShapeWrapping(Shape shape)
        {
            this.Shape = shape;
        }

        public override PPTVersion ClassVersion => PPTVersion.PPT2013;


        #region [  Code Implements  ]
        public ActionSettings ActionSettings => Shape.ActionSettings;
        public Adjustments Adjustments => Shape.Adjustments;
        public string AlternativeText { get { return Shape.AlternativeText; } set { Shape.AlternativeText = value; } }
        public AnimationSettings AnimationSettings => Shape.AnimationSettings;
        public dynamic Application => Shape.Application;
        public core.MsoAutoShapeType AutoShapeType { get { return Shape.AutoShapeType; } set { Shape.AutoShapeType = value; } }
        public core.MsoBackgroundStyleIndex BackgroundStyle { get { return Shape.BackgroundStyle; } set { Shape.BackgroundStyle = value; } }
        public core.MsoBlackWhiteMode BlackWhiteMode { get { return Shape.BlackWhiteMode; } set { Shape.BlackWhiteMode = value; } }
        public CalloutFormat Callout => Shape.Callout;
        public CanvasShapes CanvasItems => Shape.CanvasItems;
        public Chart Chart => Shape.Chart;
        public core.MsoTriState Child => Shape.Child;
        public int ConnectionSiteCount => Shape.ConnectionSiteCount;
        public core.MsoTriState Connector => Shape.Connector;
        public ConnectorFormat ConnectorFormat => Shape.ConnectorFormat;
        public int Creator => Shape.Creator;
        public CustomerData CustomerData => Shape.CustomerData;
        public Diagram Diagram => Shape.Diagram;
        public DiagramNode DiagramNode => Shape.DiagramNode;
        public FillFormat Fill => Shape.Fill;
        public core.GlowFormat Glow => Shape.Glow;
        public GroupShapes GroupItems => Shape.GroupItems;
        public core.MsoTriState HasChart => Shape.HasChart;
        public core.MsoTriState HasDiagram => Shape.HasDiagram;
        public core.MsoTriState HasDiagramNode => Shape.HasDiagramNode;
        public core.MsoTriState HasSmartArt => Shape.HasSmartArt;
        public core.MsoTriState HasTable => Shape.HasTable;
        public core.MsoTriState HasTextFrame => Shape.HasTextFrame;
        public float Height { get { return Shape.Height; } set { Shape.Height = value; } }
        public core.MsoTriState HorizontalFlip => Shape.HorizontalFlip;
        public int Id => Shape.Id;
        public float Left { get { return Shape.Left; } set { Shape.Left = value; } }
        public LineFormat Line => Shape.Line;
        public LinkFormat LinkFormat => Shape.LinkFormat;
        public core.MsoTriState LockAspectRatio { get { return Shape.LockAspectRatio; } set { Shape.LockAspectRatio = value; } }
        public MediaFormat MediaFormat => Shape.MediaFormat;
        public PpMediaType MediaType => Shape.MediaType;
        public string Name { get { return Shape.Name; } set { Shape.Name = value; } }
        public ShapeNodes Nodes => Shape.Nodes;
        public OLEFormat OLEFormat => Shape.OLEFormat;
        public dynamic Parent => Shape.Parent;
        public Shape ParentGroup => Shape.ParentGroup;
        public PictureFormat PictureFormat => Shape.PictureFormat;
        public PlaceholderFormat PlaceholderFormat => Shape.PlaceholderFormat;
        public core.ReflectionFormat Reflection => Shape.Reflection;
        public float Rotation { get { return Shape.Rotation; } set { Shape.Rotation = value; } }
        public string RTF { set { Shape.RTF = value; } }
        public core.Script Script => Shape.Script;
        public ShadowFormat Shadow => Shape.Shadow;
        public core.MsoShapeStyleIndex ShapeStyle { get { return Shape.ShapeStyle; } set { Shape.ShapeStyle = value; } }
        public core.SmartArt SmartArt => Shape.SmartArt;
        public core.SoftEdgeFormat SoftEdge => Shape.SoftEdge;
        public SoundFormat SoundFormat => Shape.SoundFormat;
        public Table Table => Shape.Table;
        public Tags Tags => Shape.Tags;
        public TextEffectFormat TextEffect => Shape.TextEffect;
        public TextFrame TextFrame => Shape.TextFrame;
        public TextFrame2 TextFrame2 => Shape.TextFrame2;
        public ThreeDFormat ThreeD => Shape.ThreeD;
        public string Title { get { return Shape.Title; } set { Shape.Title = value; } }
        public float Top { get { return Shape.Top; } set { Shape.Top = value; } }
        public core.MsoShapeType Type => Shape.Type;
        public core.MsoTriState VerticalFlip => Shape.VerticalFlip;
        public dynamic Vertices => Shape.Vertices;
        public core.MsoTriState Visible { get { return Shape.Visible; } set { Shape.Visible = value; } }
        public float Width { get { return Shape.Width; } set { Shape.Width = value; } }
        public int ZOrderPosition => Shape.ZOrderPosition;
        #endregion



        public void Apply()
        {
            Shape.Apply();
        }

        public void ApplyAnimation()
        {
            Shape.ApplyAnimation();
        }

        public void CanvasCropBottom(float Increment)
        {
            Shape.CanvasCropBottom(Increment);
        }

        public void CanvasCropLeft(float Increment)
        {
            Shape.CanvasCropLeft(Increment);
        }

        public void CanvasCropRight(float Increment)
        {
            Shape.CanvasCropRight(Increment);
        }

        public void CanvasCropTop(float Increment)
        {
            Shape.CanvasCropTop(Increment);
        }

        public void ConvertTextToSmartArt(core.SmartArtLayout Layout)
        {
            Shape.ConvertTextToSmartArt(Layout);
        }

        public void Copy()
        {
            Shape.Copy();
        }

        public void Cut()
        {
            Shape.Cut();
        }

        public void Delete()
        {
            Shape.Delete();
        }

        public ShapeRange Duplicate()
        {
            return Shape.Duplicate();
        }

        public void Export(string PathName, PpShapeFormat Filter, int ScaleWidth = 0, int ScaleHeight = 0, PpExportMode ExportMode = PpExportMode.ppRelativeToSlide)
        {
            Shape.Export(PathName, Filter, ScaleWidth, ScaleHeight, ExportMode);
        }

        public void Flip(core.MsoFlipCmd FlipCmd)
        {
            Shape.Flip(FlipCmd);
        }

        public void IncrementLeft(float Increment)
        {
            Shape.IncrementLeft(Increment);
        }

        public void IncrementRotation(float Increment)
        {
            Shape.IncrementRotation(Increment);
        }

        public void IncrementTop(float Increment)
        {
            Shape.IncrementTop(Increment);
        }

        public void PickUp()
        {
            Shape.PickUp();
        }

        public void PickupAnimation()
        {
            Shape.PickupAnimation();
        }

        public void RerouteConnections()
        {
            Shape.RerouteConnections();
        }

        public void ScaleHeight(float Factor, core.MsoTriState RelativeToOriginalSize, core.MsoScaleFrom fScale = core.MsoScaleFrom.msoScaleFromTopLeft)
        {
            Shape.ScaleHeight(Factor, RelativeToOriginalSize, fScale);
        }

        public void ScaleWidth(float Factor, core.MsoTriState RelativeToOriginalSize, core.MsoScaleFrom fScale = core.MsoScaleFrom.msoScaleFromTopLeft)
        {
            Shape.ScaleWidth(Factor, RelativeToOriginalSize, fScale);
        }

        public void Select(core.MsoTriState Replace = core.MsoTriState.msoTrue)
        {
            Shape.Select(Replace);
        }

        public void SetShapesDefaultProperties()
        {
            Shape.SetShapesDefaultProperties();
        }

        public ShapeRange Ungroup()
        {
            return Shape.Ungroup();
        }

        public void UpgradeMedia()
        {
            Shape.UpgradeMedia();
        }

        public void ZOrder(core.MsoZOrderCmd ZOrderCmd)
        {
            Shape.ZOrder(ZOrderCmd);
        }


    }
}
