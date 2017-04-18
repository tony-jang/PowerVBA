using Microsoft.Vbe.Interop;
using PowerVBA.Core.Interface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Collections;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.Connector;

namespace PowerVBA.V2010.WrapClass

{
    [Wrapped(typeof(Presentation))]
    public class PresentationWrapping : PresentationWrappingBase
    {
        public Presentation Presentation { get; }
        public PresentationWrapping(Presentation presentation)
        {
            this.Presentation = presentation;
        }

        public override PPTVersion ClassVersion => PPTVersion.PPT2010;

        public Master AddTitleMaster()
        {
            return Presentation.AddTitleMaster();
        }

        public void ApplyTemplate(string FileName)
        {
            Presentation.ApplyTemplate(FileName);
        }

        public DocumentWindow NewWindow()
        {
            return Presentation.NewWindow();
        }

        public void FollowHyperlink(string Address, string SubAddress = "", 
            bool NewWindow = false, bool AddHistory = true, string ExtraInfo = "", 
            MsoExtraInfoMethod Method = MsoExtraInfoMethod.msoMethodGet, string HeaderInfo = "")
        {
            Presentation.FollowHyperlink(Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo);
        }

        public void AddToFavorites()
        {
            Presentation.AddToFavorites();
        }

        public void Unused()
        {
            Presentation.Unused();
        }

        public void PrintOut(int From = -1, int To = -1, string PrintToFile = "", int Copies = 0, MsoTriState Collate = (MsoTriState)(-99))
        {
            Presentation.PrintOut(From, To, PrintToFile, Copies, Collate);
        }

        public void Save()
        {
            Presentation.Save();
        }

        public void SaveAs(string FileName, PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault, MsoTriState EmbedTrueTypeFonts = MsoTriState.msoTriStateMixed)
        {
            Presentation.SaveAs(FileName, FileFormat, EmbedTrueTypeFonts);
        }

        public void SaveCopyAs(string FileName, PpSaveAsFileType FileFormat = PpSaveAsFileType.ppSaveAsDefault, MsoTriState EmbedTrueTypeFonts = MsoTriState.msoTriStateMixed)
        {
            Presentation.SaveCopyAs(FileName, FileFormat, EmbedTrueTypeFonts);
        }

        public void Export(string Path, string FilterName, int ScaleWidth = 0, int ScaleHeight = 0)
        {
            Presentation.Export(Path, FilterName, ScaleWidth, ScaleHeight);
        }

        public void Close()
        {
            Presentation.Close();
        }

        public void SetUndoText(string Text)
        {
            Presentation.SetUndoText(Text);
        }

        public void UpdateLinks()
        {
            Presentation.UpdateLinks();
        }

        public void WebPagePreview()
        {
            Presentation.WebPagePreview();
        }

        public void ReloadAs(MsoEncoding cp)
        {
            Presentation.ReloadAs(cp);
        }

        public void MakeIntoTemplate(MsoTriState IsDesignTemplate)
        {
            Presentation.MakeIntoTemplate(IsDesignTemplate);
        }

        public void sblt(string s)
        {
            Presentation.sblt(s);
        }

        public void Merge(string Path)
        {
            Presentation.Merge(Path);
        }

        public void CheckIn(bool SaveChanges = true, object Comments = null, object MakePublic = null)
        {
            Presentation.CheckIn(SaveChanges, Comments, MakePublic);
        }

        public bool CanCheckIn()
        {
            return Presentation.CanCheckIn();
        }

        public void SendForReview(string Recipients = "", string Subject = "", bool ShowMessage = true, object IncludeAttachment = null)
        {
            Presentation.SendForReview(Recipients, Subject, ShowMessage, IncludeAttachment);
        }

        public void ReplyWithChanges(bool ShowMessage = true)
        {
            Presentation.ReplyWithChanges(ShowMessage);
        }

        public void EndReview()
        {
            Presentation.EndReview();
        }

        public void AddBaseline(string FileName = "")
        {
            Presentation.AddBaseline(FileName);
        }

        public void RemoveBaseline()
        {
            Presentation.RemoveBaseline();
        }

        public void SetPasswordEncryptionOptions(string PasswordEncryptionProvider, string PasswordEncryptionAlgorithm, int PasswordEncryptionKeyLength, bool PasswordEncryptionFileProperties)
        {
            Presentation.SetPasswordEncryptionOptions(PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties);
        }

        public void SendFaxOverInternet(string Recipients = "", string Subject = "", bool ShowMessage = false)
        {
            Presentation.SendFaxOverInternet(Recipients, Subject, ShowMessage);
        }

        public void NewSectionAfter(int Index, bool AfterSlide, string sectionTitle, out int newSectionIndex)
        {
            Presentation.NewSectionAfter(Index, AfterSlide, sectionTitle,out newSectionIndex);
        }

        public void DeleteSection(int Index)
        {
            Presentation.DeleteSection(Index);
        }

        public void DisableSections()
        {
            Presentation.DisableSections();
        }

        public string sectionTitle(int Index)
        {
            return Presentation.sectionTitle(Index);
        }

        public void RemoveDocumentInformation(PpRemoveDocInfoType Type)
        {
            Presentation.RemoveDocumentInformation(Type);
        }

        public void CheckInWithVersion(bool SaveChanges = true, object Comments = null, object MakePublic = null, object VersionType = null)
        {
            Presentation.CheckInWithVersion(SaveChanges, Comments, MakePublic, VersionType);
        }

        public void ExportAsFixedFormat(string Path, PpFixedFormatType FixedFormatType, PpFixedFormatIntent Intent = PpFixedFormatIntent.ppFixedFormatIntentScreen, MsoTriState FrameSlides = MsoTriState.msoFalse, PpPrintHandoutOrder HandoutOrder = PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType OutputType = PpPrintOutputType.ppPrintOutputSlides, MsoTriState PrintHiddenSlides = MsoTriState.msoFalse, PrintRange PrintRange = null, PpPrintRangeType RangeType = PpPrintRangeType.ppPrintAll, string SlideShowName = "", bool IncludeDocProperties = false, bool KeepIRMSettings = true, bool DocStructureTags = true, bool BitmapMissingFonts = true, bool UseISO19005_1 = false, object ExternalExporter = null)
        {
            Presentation.ExportAsFixedFormat(Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, 
                                             OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, 
                                             KeepIRMSettings, DocStructureTags, BitmapMissingFonts, UseISO19005_1, ExternalExporter);
        }

        public WorkflowTasks GetWorkflowTasks()
        {
            return Presentation.GetWorkflowTasks();
        }

        public WorkflowTemplates GetWorkflowTemplates()
        {
            return Presentation.GetWorkflowTemplates();
        }

        public void LockServerFile()
        {
            Presentation.LockServerFile();
        }

        public void ApplyTheme(string themeName)
        {
            Presentation.ApplyTheme(themeName);
        }

        public void PublishSlides(string SlideLibraryUrl, bool Overwrite = false, bool UseSlideOrder = false)
        {
            Presentation.PublishSlides(SlideLibraryUrl, Overwrite, UseSlideOrder);
        }

        public void Convert()
        {
            Presentation.Convert();
        }

        public void MergeWithBaseline(string withPresentation, string baselinePresentation)
        {
            Presentation.MergeWithBaseline(withPresentation, baselinePresentation);
        }

        public void AcceptAll()
        {
            Presentation.AcceptAll();
        }

        public void RejectAll()
        {
            Presentation.RejectAll();
        }

        public void EnsureAllMediaUpgraded()
        {
            Presentation.EnsureAllMediaUpgraded();
        }

        public void Convert2(string FileName)
        {
            Presentation.Convert2(FileName);
        }

        public void CreateVideo(string FileName, bool UseTimingsAndNarrations = true, int DefaultSlideDuration = 5, int VertResolution = 720, int FramesPerSecond = 30, int Quality = 85)
        {
            Presentation.CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality);
        }

        public Microsoft.Office.Interop.PowerPoint.Application Application => Presentation.Application;
        public dynamic Parent => Presentation.Parent;
        public Master SlideMaster => Presentation.SlideMaster;
        public Master TitleMaster => Presentation.TitleMaster;
        public MsoTriState HasTitleMaster => Presentation.HasTitleMaster;
        public string TemplateName => Presentation.TemplateName;
        public Master NotesMaster => Presentation.NotesMaster;
        public Master HandoutMaster => Presentation.HandoutMaster;
        public Slides Slides => Presentation.Slides;
        public PageSetup PageSetup => Presentation.PageSetup;
        public ColorSchemes ColorSchemes => Presentation.ColorSchemes;
        public ExtraColors ExtraColors => Presentation.ExtraColors;
        public SlideShowSettings SlideShowSettings => Presentation.SlideShowSettings;
        public Fonts Fonts => Presentation.Fonts;
        public DocumentWindows Windows => Presentation.Windows;
        public Tags Tags => Presentation.Tags;
        public Microsoft.Office.Interop.PowerPoint.Shape DefaultShape => Presentation.DefaultShape;
        public dynamic BuiltInDocumentProperties => Presentation.BuiltInDocumentProperties;
        public dynamic CustomDocumentProperties => Presentation.CustomDocumentProperties;
        public VBProject VBProject => Presentation.VBProject;
        public MsoTriState ReadOnly => Presentation.ReadOnly;
        public string FullName => Presentation.FullName;
        public string Name => Presentation.Name;
        public string Path => Presentation.Path;
        public MsoTriState Saved { set { Presentation.Saved = value; } get { return Presentation.Saved; } }
        public PpDirection LayoutDirection { set { Presentation.LayoutDirection = value; } get { return Presentation.LayoutDirection; } }
        public PrintOptions PrintOptions => Presentation.PrintOptions;
        public dynamic Container => Presentation.Container;
        public MsoTriState DisplayComments { set { Presentation.DisplayComments = value; } get { return Presentation.DisplayComments; } }
        public PpFarEastLineBreakLevel FarEastLineBreakLevel { set { Presentation.FarEastLineBreakLevel = value; } get { return Presentation.FarEastLineBreakLevel; } }
        public string NoLineBreakBefore { set { Presentation.NoLineBreakBefore = value; } get { return Presentation.NoLineBreakBefore; } }
        public string NoLineBreakAfter { set { Presentation.NoLineBreakAfter = value; } get { return Presentation.NoLineBreakAfter; } }
        public SlideShowWindow SlideShowWindow => Presentation.SlideShowWindow;
        public MsoFarEastLineBreakLanguageID FarEastLineBreakLanguage { set { Presentation.FarEastLineBreakLanguage = value; } get { return Presentation.FarEastLineBreakLanguage; } }
        public MsoLanguageID DefaultLanguageID { set { Presentation.DefaultLanguageID = value; } get { return Presentation.DefaultLanguageID; } }
        public CommandBars CommandBars => Presentation.CommandBars;
        public PublishObjects PublishObjects => Presentation.PublishObjects;
        public WebOptions WebOptions => Presentation.WebOptions;
        public HTMLProject HTMLProject => Presentation.HTMLProject;
        public MsoTriState EnvelopeVisible { set { Presentation.EnvelopeVisible = value; } get { return Presentation.EnvelopeVisible; } }
        public MsoTriState VBASigned => Presentation.VBASigned;
        public MsoTriState SnapToGrid { set { Presentation.SnapToGrid = value; } get { return Presentation.SnapToGrid; } }
        public float GridDistance { set { Presentation.GridDistance = value; } get { return Presentation.GridDistance; } }
        public Designs Designs => Presentation.Designs;
        public SignatureSet Signatures => Presentation.Signatures;
        public MsoTriState RemovePersonalInformation { set { Presentation.RemovePersonalInformation = value; } get { return Presentation.RemovePersonalInformation; } }
        public PpRevisionInfo HasRevisionInfo => Presentation.HasRevisionInfo;
        public string PasswordEncryptionProvider => Presentation.PasswordEncryptionProvider;
        public string PasswordEncryptionAlgorithm => Presentation.PasswordEncryptionAlgorithm;
        public int PasswordEncryptionKeyLength => Presentation.PasswordEncryptionKeyLength;
        public bool PasswordEncryptionFileProperties => Presentation.PasswordEncryptionFileProperties;
        public string Password { set { Presentation.Password = value; } get { return Presentation.Password; } }
        public string WritePassword { set { Presentation.WritePassword = value; } get { return Presentation.WritePassword; } }
        public Permission Permission => Presentation.Permission;
        public SharedWorkspace SharedWorkspace => Presentation.SharedWorkspace;
        public Sync Sync => Presentation.Sync;
        public DocumentLibraryVersions DocumentLibraryVersions => Presentation.DocumentLibraryVersions;
        public MetaProperties ContentTypeProperties => Presentation.ContentTypeProperties;
        public int SectionCount => Presentation.SectionCount;
        public bool HasSections => Presentation.HasSections;
        public ServerPolicy ServerPolicy => Presentation.ServerPolicy;
        public DocumentInspectors DocumentInspectors => Presentation.DocumentInspectors;
        public bool HasVBProject => Presentation.HasVBProject;
        public CustomXMLParts CustomXMLParts => Presentation.CustomXMLParts;
        public bool Final { set { Presentation.Final = value; } get { return Presentation.Final; } }
        public CustomerData CustomerData => Presentation.CustomerData;
        public Research Research => Presentation.Research;
        public string EncryptionProvider { set { Presentation.EncryptionProvider = value; } get { return Presentation.EncryptionProvider; } }
        public SectionProperties SectionProperties => Presentation.SectionProperties;
        public Coauthoring Coauthoring => Presentation.Coauthoring;
        public bool InMergeMode => Presentation.InMergeMode;
        public Broadcast Broadcast => Presentation.Broadcast;
        public bool HasNotesMaster => Presentation.HasNotesMaster;
        public bool HasHandoutMaster => Presentation.HasHandoutMaster;
        public PpMediaTaskStatus CreateVideoStatus => Presentation.CreateVideoStatus;
    }
}
