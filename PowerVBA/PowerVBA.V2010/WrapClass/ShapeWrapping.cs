using Microsoft.Vbe.Interop;
using PowerVBA.Core.Interface;
using Microsoft.Office.Interop.PowerPoint;
using System.Collections;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.Connector;
using System;
using c=Microsoft.Office.Core;
using System.Windows.Media;

namespace PowerVBA.V2010.WrapClass

{
    [Wrapped(typeof(Shape))]
    public class ShapeWrapping : ShapeWrappingBase
    {
        public Shape Shape { get; }
        public ShapeWrapping(Shape shape)
        {
            this.Shape = shape;
        }

        public override PPTVersion ClassVersion => PPTVersion.PPT2010;

        public void Apply()
        {
            Shape.Apply();
        }

        public void Delete()
        {
            Shape.Delete();
        }

        public void Flip(c.MsoFlipCmd FlipCmd)
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

        public void RerouteConnections()
        {
            Shape.RerouteConnections();
        }

        public void ScaleHeight(float Factor, c.MsoTriState RelativeToOriginalSize, c.MsoScaleFrom fScale = c.MsoScaleFrom.msoScaleFromTopLeft)
        {
            Shape.ScaleHeight(Factor, RelativeToOriginalSize, fScale);
        }

        public void ScaleWidth(float Factor, c.MsoTriState RelativeToOriginalSize, c.MsoScaleFrom fScale = c.MsoScaleFrom.msoScaleFromTopLeft)
        {
            Shape.ScaleWidth(Factor, RelativeToOriginalSize, fScale);
        }

        public void SetShapesDefaultProperties()
        {
            Shape.SetShapesDefaultProperties();
        }

        public ShapeRange Ungroup()
        {
            return Shape.Ungroup();
        }

        public void ZOrder(c.MsoZOrderCmd ZOrderCmd)
        {
            Shape.ZOrder(ZOrderCmd);
        }

        public void Cut()
        {
            Shape.Cut();
        }

        public void Copy()
        {
            Shape.Copy();
        }

        public void Select(c.MsoTriState Replace = c.MsoTriState.msoTrue)
        {
            Shape.Select(Replace);
        }

        public ShapeRange Duplicate()
        {
            return Shape.Duplicate();
        }

        public void Export(string PathName, PpShapeFormat Filter, int ScaleWidth = 0, int ScaleHeight = 0, PpExportMode ExportMode = PpExportMode.ppRelativeToSlide)
        {
            Shape.Export(PathName, Filter, ScaleWidth, ScaleHeight, ExportMode);
        }

        public void CanvasCropLeft(float Increment)
        {
            Shape.CanvasCropLeft(Increment);
        }

        public void CanvasCropTop(float Increment)
        {
            Shape.CanvasCropTop(Increment);
        }

        public void CanvasCropRight(float Increment)
        {
            Shape.CanvasCropRight(Increment);
        }

        public void CanvasCropBottom(float Increment)
        {
            Shape.CanvasCropBottom(Increment);
        }

        public void ConvertTextToSmartArt(c.SmartArtLayout Layout)
        {
            Shape.ConvertTextToSmartArt(Layout);
        }

        public void PickupAnimation()
        {
            Shape.PickupAnimation();
        }

        public void ApplyAnimation()
        {
            Shape.ApplyAnimation();
        }

        public void UpgradeMedia()
        {
            Shape.UpgradeMedia();
        }

        public dynamic Application => Shape.Application;
        public int Creator => Shape.Creator;
        public dynamic Parent => Shape.Parent;
        public Adjustments Adjustments => Shape.Adjustments;
        public c.MsoAutoShapeType AutoShapeType { set { Shape.AutoShapeType = value; } get { return Shape.AutoShapeType; } }
        public c.MsoBlackWhiteMode BlackWhiteMode { set { Shape.BlackWhiteMode = value; } get { return Shape.BlackWhiteMode; } }
        public CalloutFormat Callout => Shape.Callout;
        public int ConnectionSiteCount => Shape.ConnectionSiteCount;
        public c.MsoTriState Connector => Shape.Connector;
        public ConnectorFormat ConnectorFormat => Shape.ConnectorFormat;
        public FillFormat Fill => Shape.Fill;
        public GroupShapes GroupItems => Shape.GroupItems;
        public override float Height { set { Shape.Height = value; } get { return Shape.Height; } }
        public c.MsoTriState HorizontalFlip => Shape.HorizontalFlip;
        public override float Left { set { Shape.Left = value; } get { return Shape.Left; } }
        public LineFormat Line => Shape.Line;
        public c.MsoTriState LockAspectRatio { set { Shape.LockAspectRatio = value; } get { return Shape.LockAspectRatio; } }
        public override string Name { set { Shape.Name = value; } get { return Shape.Name; } }
        public ShapeNodes Nodes => Shape.Nodes;
        public float Rotation { set { Shape.Rotation = value; } get { return Shape.Rotation; } }
        public PictureFormat PictureFormat => Shape.PictureFormat;
        public ShadowFormat Shadow => Shape.Shadow;
        public TextEffectFormat TextEffect => Shape.TextEffect;
        public TextFrame TextFrame => Shape.TextFrame;
        public ThreeDFormat ThreeD => Shape.ThreeD;
        public override float Top { set { Shape.Top = value; } get { return Shape.Top; } }
        public c.MsoShapeType Type => Shape.Type;
        public c.MsoTriState VerticalFlip => Shape.VerticalFlip;
        public dynamic Vertices => Shape.Vertices;
        public c.MsoTriState Visible { set { Shape.Visible = value; } get { return Shape.Visible; } }
        public override float Width { set { Shape.Width = value; } get { return Shape.Width; } }
        public int ZOrderPosition => Shape.ZOrderPosition;
        public OLEFormat OLEFormat => Shape.OLEFormat;
        public LinkFormat LinkFormat => Shape.LinkFormat;
        public PlaceholderFormat PlaceholderFormat => Shape.PlaceholderFormat;
        public AnimationSettings AnimationSettings => Shape.AnimationSettings;
        public ActionSettings ActionSettings => Shape.ActionSettings;
        public Tags Tags => Shape.Tags;
        public PpMediaType MediaType => Shape.MediaType;
        public c.MsoTriState HasTextFrame => Shape.HasTextFrame;
        public SoundFormat SoundFormat => Shape.SoundFormat;
        public c.Script Script => Shape.Script;
        public string AlternativeText { set { Shape.AlternativeText = value; } get { return Shape.AlternativeText; } }
        public c.MsoTriState HasTable => Shape.HasTable;
        public Table Table => Shape.Table;
        public c.MsoTriState HasDiagram => Shape.HasDiagram;
        public Diagram Diagram => Shape.Diagram;
        public c.MsoTriState HasDiagramNode => Shape.HasDiagramNode;
        public DiagramNode DiagramNode => Shape.DiagramNode;
        public c.MsoTriState Child => Shape.Child;
        public Shape ParentGroup => Shape.ParentGroup;
        public CanvasShapes CanvasItems => Shape.CanvasItems;
        public int Id => Shape.Id;
        public string RTF { set { Shape.RTF = value; } }
        public CustomerData CustomerData => Shape.CustomerData;
        public TextFrame2 TextFrame2 => Shape.TextFrame2;
        public c.MsoTriState HasChart => Shape.HasChart;
        public c.MsoShapeStyleIndex ShapeStyle { set { Shape.ShapeStyle = value; } get { return Shape.ShapeStyle; } }
        public c.MsoBackgroundStyleIndex BackgroundStyle { set { Shape.BackgroundStyle = value; } get { return Shape.BackgroundStyle; } }
        public c.SoftEdgeFormat SoftEdge => Shape.SoftEdge;
        public c.GlowFormat Glow => Shape.Glow;
        public c.ReflectionFormat Reflection => Shape.Reflection;
        public Chart Chart => Shape.Chart;
        public c.MsoTriState HasSmartArt => Shape.HasSmartArt;
        public c.SmartArt SmartArt => Shape.SmartArt;
        public string Title { set { Shape.Title = value; } get { return Shape.Title; } }
        public MediaFormat MediaFormat => Shape.MediaFormat;

        public override string ShapeType => Shape.Type.ToString();

        public override Color RGB
        {
            get
            {
                int colorInt = Shape.Fill.BackColor.RGB;
                var b = (byte)((colorInt >> 16) & 0xff);
                var g = (byte)((colorInt >> 8) & 0xff);
                var r = (byte)(colorInt & 0xff);

                return Color.FromRgb(r, g, b);
            }
        }

        public override Color ForeRGB
        {
            get
            {
                int colorInt = Shape.Fill.ForeColor.RGB;
                var b = (byte)((colorInt >> 16) & 0xff);
                var g = (byte)((colorInt >> 8) & 0xff);
                var r = (byte)(colorInt & 0xff);

                return Color.FromRgb(r, g, b);
            }
        }

        public override void Delete(out bool success)
        {
            try
            {
                Delete();
                success = true;
            }
            catch (Exception)
            {
                success = false;
            }

        }

    }
}
