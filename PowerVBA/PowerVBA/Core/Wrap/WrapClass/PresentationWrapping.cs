using Microsoft.Office.Interop.PowerPoint;
using core = Microsoft.Office.Core;
using PowerVBA.Core.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace PowerVBA.Core.Wrap.WrapClass
{
    [Wrapped(typeof(Presentation))]
    class PresentationWrapping : IWrappingClass
    {
        public Presentation Presentation { get; }

        public PresentationWrapping(Presentation presentation)
        {
            this.Presentation = presentation;
        }


        public Microsoft.Office.Interop.PowerPoint.Application Application => Presentation.Application;
        public Broadcast Broadcast => Presentation.Broadcast;
        public dynamic BuiltInDocumentProperties => Presentation.BuiltInDocumentProperties;
        public bool ChartDataPointTrack { set { Presentation.ChartDataPointTrack = value; } get { return Presentation.ChartDataPointTrack; } }
        public Coauthoring Coauthoring => Presentation.Coauthoring;
        public ColorSchemes ColorSchemes => Presentation.ColorSchemes;
        public core.CommandBars CommandBars => Presentation.CommandBars;
        public dynamic Container => Presentation.Container;
        public core.MetaProperties ContentTypeProperties => Presentation.ContentTypeProperties;
        public PpMediaTaskStatus CreateVideoStatus => Presentation.CreateVideoStatus;
        public dynamic CustomDocumentProperties => Presentation.CustomDocumentProperties;
        public CustomerData CustomerData => Presentation.CustomerData;
        public core.CustomXMLParts CustomXMLParts => Presentation.CustomXMLParts;
        public core.MsoLanguageID DefaultLanguageID { set { Presentation.DefaultLanguageID = value; } get { return Presentation.DefaultLanguageID; } }
        public Shape DefaultShape => Presentation.DefaultShape;
        public Designs Designs => Presentation.Designs;
        public core.MsoTriState DisplayComments { set { Presentation.DisplayComments = value; } get { return Presentation.DisplayComments; } }
        public core.DocumentInspectors DocumentInspectors => Presentation.DocumentInspectors;
        public core.DocumentLibraryVersions DocumentLibraryVersions => Presentation.DocumentLibraryVersions;
        public string EncryptionProvider { set { Presentation.EncryptionProvider = value; } get { return Presentation.EncryptionProvider; } }
        public core.MsoTriState EnvelopeVisible { set { Presentation.EnvelopeVisible = value; } get { return Presentation.EnvelopeVisible; } }
        public ExtraColors ExtraColors => Presentation.ExtraColors;
        public core.MsoFarEastLineBreakLanguageID FarEastLineBreakLanguage { set { Presentation.FarEastLineBreakLanguage = value; } get { return Presentation.FarEastLineBreakLanguage; } }
        public PpFarEastLineBreakLevel FarEastLineBreakLevel { set { Presentation.FarEastLineBreakLevel = value; } get { return Presentation.FarEastLineBreakLevel; } }
        public bool Final { set { Presentation.Final = value; } get { return Presentation.Final; } }
        public Fonts Fonts => Presentation.Fonts;
        public string FullName => Presentation.FullName;
        public float GridDistance { set { Presentation.GridDistance = value; } get { return Presentation.GridDistance; } }
        public Guides Guides => Presentation.Guides;
        public Master HandoutMaster => Presentation.HandoutMaster;
        public bool HasHandoutMaster => Presentation.HasHandoutMaster;
        public bool HasNotesMaster => Presentation.HasNotesMaster;
        public PpRevisionInfo HasRevisionInfo => Presentation.HasRevisionInfo;
        public bool HasSections => Presentation.HasSections;
        public core.MsoTriState HasTitleMaster => Presentation.HasTitleMaster;
        public bool HasVBProject => Presentation.HasVBProject;
        public core.HTMLProject HTMLProject => Presentation.HTMLProject;
        public bool InMergeMode => Presentation.InMergeMode;
        public PpDirection LayoutDirection { set { Presentation.LayoutDirection = value; } get { return Presentation.LayoutDirection; } }
        public string Name => Presentation.Name;
        public string NoLineBreakAfter { set { Presentation.NoLineBreakAfter = value; } get { return Presentation.NoLineBreakAfter; } }
        public string NoLineBreakBefore { set { Presentation.NoLineBreakBefore = value; } get { return Presentation.NoLineBreakBefore; } }
        public Master NotesMaster => Presentation.NotesMaster;
        public PageSetup PageSetup => Presentation.PageSetup;
        public dynamic Parent => Presentation.Parent;
        public string Password { set { Presentation.Password = value; } get { return Presentation.Password; } }
        public string PasswordEncryptionAlgorithm => Presentation.PasswordEncryptionAlgorithm;
        public bool PasswordEncryptionFileProperties => Presentation.PasswordEncryptionFileProperties;
        public int PasswordEncryptionKeyLength => Presentation.PasswordEncryptionKeyLength;
        public string PasswordEncryptionProvider => Presentation.PasswordEncryptionProvider;
        public string Path => Presentation.Path;
        public core.Permission Permission => Presentation.Permission;
        public PrintOptions PrintOptions => Presentation.PrintOptions;
        public PublishObjects PublishObjects => Presentation.PublishObjects;
        public core.MsoTriState ReadOnly => Presentation.ReadOnly;
        public core.MsoTriState RemovePersonalInformation { set { Presentation.RemovePersonalInformation = value; } get { return Presentation.RemovePersonalInformation; } }
        public Research Research => Presentation.Research;
        public core.MsoTriState Saved { set { Presentation.Saved = value; } get { return Presentation.Saved; } }
        public int SectionCount => Presentation.SectionCount;
        public SectionProperties SectionProperties => Presentation.SectionProperties;
        public core.ServerPolicy ServerPolicy => Presentation.ServerPolicy;
        public core.SharedWorkspace SharedWorkspace => Presentation.SharedWorkspace;
        public core.SignatureSet Signatures => Presentation.Signatures;
        public Master SlideMaster => Presentation.SlideMaster;
        public Slides Slides => Presentation.Slides;
        public SlideShowSettings SlideShowSettings => Presentation.SlideShowSettings;
        public SlideShowWindow SlideShowWindow => Presentation.SlideShowWindow;
        public core.MsoTriState SnapToGrid { set { Presentation.SnapToGrid = value; } get { return Presentation.SnapToGrid; } }
        public core.Sync Sync => Presentation.Sync;
        public Tags Tags => Presentation.Tags;
        public string TemplateName => Presentation.TemplateName;
        public Master TitleMaster => Presentation.TitleMaster;
        public core.MsoTriState VBASigned => Presentation.VBASigned;
        public VBProject VBProject => Presentation.VBProject;
        public WebOptions WebOptions => Presentation.WebOptions;
        public DocumentWindows Windows => Presentation.Windows;
        public string WritePassword { set { Presentation.WritePassword = value; } get { return Presentation.WritePassword; } }
        public void AcceptAll()
        {
            Presentation.AcceptAll();
        }

        public void AddBaseline(string FileName = "")
        {
            Presentation.AddBaseline(FileName);
        }

        public Master AddTitleMaster()
        {
            return Presentation.AddTitleMaster();
        }

        public void AddToFavorites()
        {
            Presentation.AddToFavorites();
        }

        public void ApplyTemplate(string FileName)
        {
            Presentation.ApplyTemplate(FileName);
        }

        public void ApplyTemplate2(string FileName, string VariantGUID)
        {
            Presentation.ApplyTemplate2(FileName, VariantGUID);
        }

        public void ApplyTheme(string themeName)
        {
            Presentation.ApplyTheme(themeName);
        }

        public bool CanCheckIn()
        {
            return Presentation.CanCheckIn();
        }

        public void CheckIn(bool SaveChanges = false, object Comments = null, object MakePublic = null)
        {
            Presentation.CheckIn(SaveChanges, Comments, MakePublic);
        }

        public void CheckInWithVersion(bool SaveChanges = false, object Comments = null, object MakePublic = null, object VersionType = null)
        {
            Presentation.CheckInWithVersion(SaveChanges, Comments, MakePublic, VersionType);
        }

        public void Close()
        {
            Presentation.Close();
        }

        public void Convert()
        {
            Presentation.Convert();
        }

        public void Convert2(string FileName)
        {
            Presentation.Convert2(FileName);
        }

        public void CreateVideo(string FileName, bool UseTimingsAndNarrations = false, int DefaultSlideDuration = 5, int VertResolution = 720, int FramesPerSecond = 30, int Quality = 85)
        {
            Presentation.CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality);
        }

        public void DeleteSection(int Index)
        {
            Presentation.DeleteSection(Index);
        }

        public void DisableSections()
        {
            Presentation.DisableSections();
        }

        public void EndReview()
        {
            Presentation.EndReview();
        }

        public void EnsureAllMediaUpgraded()
        {
            Presentation.EnsureAllMediaUpgraded();
        }

        public void Export(string Path, string FilterName, int ScaleWidth = 0, int ScaleHeight = 0)
        {
            Presentation.Export(Path, FilterName, ScaleWidth, ScaleHeight);
        }

        public void ExportAsFixedFormat(string Path, PpFixedFormatType FixedFormatType, 
            PpFixedFormatIntent Intent = PpFixedFormatIntent.ppFixedFormatIntentScreen,
            core.MsoTriState FrameSlides = core.MsoTriState.msoFalse, 
            PpPrintHandoutOrder HandoutOrder = PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, 
            PpPrintOutputType OutputType = PpPrintOutputType.ppPrintOutputSlides,
            core.MsoTriState PrintHiddenSlides = core.MsoTriState.msoFalse, 
            PrintRange PrintRange = null, PpPrintRangeType RangeType = PpPrintRangeType.ppPrintAll, 
            string SlideShowName = "", bool IncludeDocProperties = false, bool KeepIRMSettings = false, 
            bool DocStructureTags = false, bool BitmapMissingFonts = false, bool UseISO19005_1 = false, object ExternalExporter = null)
        {
            Presentation.ExportAsFixedFormat(Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, ExternalExporter);
        }

        public void ExportAsFixedFormat2(string Path, PpFixedFormatType FixedFormatType, PpFixedFormatIntent Intent = PpFixedFormatIntent.ppFixedFormatIntentScreen, core.MsoTriState FrameSlides = core.MsoTriState.msoFalse, PpPrintHandoutOrder HandoutOrder = PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType OutputType = PpPrintOutputType.ppPrintOutputSlides, core.MsoTriState PrintHiddenSlides = core.MsoTriState.msoFalse, PrintRange PrintRange = null, PpPrintRangeType RangeType = PpPrintRangeType.ppPrintAll, string SlideShowName = "", bool IncludeDocProperties = false, bool KeepIRMSettings = false, bool DocStructureTags = false, bool BitmapMissingFonts = false, bool UseISO19005_1 = false, bool IncludeMarkup = false, object ExternalExporter = null)
        {
            Presentation.ExportAsFixedFormat2(Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, IncludeMarkup, ExternalExporter);
        }

        public void FollowHyperlink(string Address, string SubAddress = "", bool NewWindow = false, bool AddHistory = false, 
            string ExtraInfo = "", core.MsoExtraInfoMethod Method = core.MsoExtraInfoMethod.msoMethodGet, string HeaderInfo = "")
        {
            Presentation.FollowHyperlink(Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo);
        }

        public core.WorkflowTasks GetWorkflowTasks()
        {
            return Presentation.GetWorkflowTasks();
        }

        public core.WorkflowTemplates GetWorkflowTemplates()
        {
            return Presentation.GetWorkflowTemplates();
        }

        public void LockServerFile()
        {
            Presentation.LockServerFile();
        }

        public void MakeIntoTemplate(core.MsoTriState IsDesignTemplate)
        {
            Presentation.MakeIntoTemplate(IsDesignTemplate);
        }

        public void Merge(string Path)
        {
            Presentation.Merge(Path);
        }

        public void MergeWithBaseline(string withPresentation, string baselinePresentation)
        {
            Presentation.MergeWithBaseline(withPresentation, baselinePresentation);
        }

        public void NewSectionAfter(int Index, bool AfterSlide, string sectionTitle, out int newSectionIndex)
        {
            Presentation.NewSectionAfter(Index, AfterSlide, sectionTitle, out newSectionIndex);
        }

        public DocumentWindow NewWindow()
        {
            return Presentation.NewWindow();
        }

        public void PrintOut(int From = -1, int To = -1, string PrintToFile = "", int Copies = 0, core.MsoTriState Collate = (core.MsoTriState)(-99))
        {
            Presentation.PrintOut(From, To, PrintToFile, Copies, Collate);
        }

        public void PublishSlides(string SlideLibraryUrl, bool Overwrite = false, bool UseSlideOrder = false)
        {
            Presentation.PublishSlides(SlideLibraryUrl, Overwrite, UseSlideOrder);
        }

        public void RejectAll()
        {
            Presentation.RejectAll();
        }

        public void ReloadAs(core.MsoEncoding cp)
        {
            Presentation.ReloadAs(cp);
        }

        public void RemoveBaseline()
        {
            Presentation.RemoveBaseline();
        }

        public void RemoveDocumentInformation(PpRemoveDocInfoType Type)
        {
            Presentation.RemoveDocumentInformation(Type);
        }

        public void ReplyWithChanges(bool ShowMessage = false)
        {
            Presentation.ReplyWithChanges(ShowMessage);
        }

        public void Save()
        {
            Presentation.Save();
        }

        public void SaveAs(string FileName, PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault,
            core.MsoTriState EmbedTrueTypeFonts = core.MsoTriState.msoTriStateMixed)
        {
            Presentation.SaveAs(FileName, FileFormat, EmbedTrueTypeFonts);
        }

        public void SaveCopyAs(string FileName, PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault,
            core.MsoTriState EmbedTrueTypeFonts = core.MsoTriState.msoTriStateMixed)
        {
            Presentation.SaveCopyAs(FileName, FileFormat,EmbedTrueTypeFonts);
        }

        public void sblt(string s)
        {
            Presentation.sblt(s);
        }

        public string sectionTitle(int Index)
        {
            return Presentation.sectionTitle(Index);
        }

        public void SendFaxOverInternet(string Recipients = "", string Subject = "", bool ShowMessage = false)
        {
            Presentation.SendFaxOverInternet(Recipients, Subject, ShowMessage);
        }

        public void SendForReview(string Recipients = "", string Subject = "", bool ShowMessage = false, object IncludeAttachment = null)
        {
            Presentation.SendForReview(Recipients, Subject, ShowMessage, IncludeAttachment);
        }

        public void SetPasswordEncryptionOptions(string PasswordEncryptionProvider, string PasswordEncryptionAlgorithm, int PasswordEncryptionKeyLength, bool PasswordEncryptionFileProperties)
        {
            Presentation.SetPasswordEncryptionOptions(PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties);
        }

        public void SetUndoText(string Text)
        {
            Presentation.SetUndoText(Text);
        }

        public void Unused()
        {
            Presentation.Unused();
        }

        public void UpdateLinks()
        {
            Presentation.UpdateLinks();
        }

        public void WebPagePreview()
        {
            Presentation.WebPagePreview();
        }


    }
}
