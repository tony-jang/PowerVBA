using Microsoft.Vbe.Interop;
using ppt=Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Collections;
using PowerVBA.Core.Interface;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.Connector;

namespace PowerVBA.V2013.Wrap.WrapClass

{
    [Wrapped(typeof(ppt._Application))]
    public class ApplicationWrapping : ApplicationWrappingBase
	{
		public ppt._Application Application { get; }
        public ApplicationWrapping(ppt._Application application)
        {
            this.Application = application;
        }
        public override PPTVersion ClassVersion => PPTVersion.PPT2013;

        public MsoTriState Active => Application.Active;
        public int ActiveEncryptionSession => Application.ActiveEncryptionSession;
        public ppt.Presentation ActivePresentation => Application.ActivePresentation;
        public string ActivePrinter => Application.ActivePrinter;
        public ppt.ProtectedViewWindow ActiveProtectedViewWindow => Application.ActiveProtectedViewWindow;
        public ppt.DocumentWindow ActiveWindow => Application.ActiveWindow;
        public ppt.AddIns AddIns => Application.AddIns;
        public AnswerWizard AnswerWizard => Application.AnswerWizard;
        public IAssistance Assistance => Application.Assistance;
        public Assistant Assistant => Application.Assistant;
        public ppt.AutoCorrect AutoCorrect => Application.AutoCorrect;
        public MsoAutomationSecurity AutomationSecurity { set { Application.AutomationSecurity = value; } get { return Application.AutomationSecurity; } }
        public string Build => Application.Build;
        public string Caption { set { Application.Caption = value; } get { return Application.Caption; } }
        public bool ChartDataPointTrack { set { Application.ChartDataPointTrack = value; } get { return Application.ChartDataPointTrack; } }
        public COMAddIns COMAddIns => Application.COMAddIns;
        public CommandBars CommandBars => Application.CommandBars;
        public int Creator => Application.Creator;
        public ppt.DefaultWebOptions DefaultWebOptions => Application.DefaultWebOptions;
        public dynamic Dialogs => Application.Dialogs;
        public ppt.PpAlertLevel DisplayAlerts { set { Application.DisplayAlerts = value; } get { return Application.DisplayAlerts; } }
        public bool DisplayDocumentInformationPanel { set { Application.DisplayDocumentInformationPanel = value; } get { return Application.DisplayDocumentInformationPanel; } }
        public MsoTriState DisplayGridLines { set { Application.DisplayGridLines = value; } get { return Application.DisplayGridLines; } }
        public MsoTriState DisplayGuides { set { Application.DisplayGuides = value; } get { return Application.DisplayGuides; } }
        public MsoFeatureInstall FeatureInstall { set { Application.FeatureInstall = value; } get { return Application.FeatureInstall; } }
        public ppt.FileConverters FileConverters => Application.FileConverters;
        public IFind FileFind => Application.FileFind;
        public FileSearch FileSearch => Application.FileSearch;
        public MsoFileValidationMode FileValidation { set { Application.FileValidation = value; } get { return Application.FileValidation; } }
        public float Height { set { Application.Height = value; } get { return Application.Height; } }
        public int HWND => Application.HWND;
        public bool IsSandboxed => Application.IsSandboxed;
        public LanguageSettings LanguageSettings => Application.LanguageSettings;
        public float Left { set { Application.Left = value; } get { return Application.Left; } }
        public dynamic Marker => Application.Marker;
        public MsoDebugOptions MsoDebugOptions => Application.MsoDebugOptions;
        public string Name => Application.Name;
        public NewFile NewPresentation => Application.NewPresentation;
        public string OperatingSystem => Application.OperatingSystem;
        public ppt.Options Options => Application.Options;
        public string Path => Application.Path;
        public ppt.Presentations Presentations => Application.Presentations;
        public string ProductCode => Application.ProductCode;
        public ppt.ProtectedViewWindows ProtectedViewWindows => Application.ProtectedViewWindows;
        public ppt.ResampleMediaTasks ResampleMediaTasks => Application.ResampleMediaTasks;
        public MsoTriState ShowStartupDialog { set { Application.ShowStartupDialog = value; } get { return Application.ShowStartupDialog; } }
        public MsoTriState ShowWindowsInTaskbar { set { Application.ShowWindowsInTaskbar = value; } get { return Application.ShowWindowsInTaskbar; } }
        public ppt.SlideShowWindows SlideShowWindows => Application.SlideShowWindows;
        public SmartArtColors SmartArtColors => Application.SmartArtColors;
        public SmartArtLayouts SmartArtLayouts => Application.SmartArtLayouts;
        public SmartArtQuickStyles SmartArtQuickStyles => Application.SmartArtQuickStyles;
        public float Top { set { Application.Top = value; } get { return Application.Top; } }
        public VBE VBE => Application.VBE;
        public string Version => Application.Version;
        public MsoTriState Visible { set { Application.Visible = value; } get { return Application.Visible; } }
        public float Width { set { Application.Width = value; } get { return Application.Width; } }
        public ppt.DocumentWindows Windows => Application.Windows;
        public ppt.PpWindowState WindowState { set { Application.WindowState = value; } get { return Application.WindowState; } }
        public void Activate()
        {
            Application.Activate();
        }

        public bool GetOptionFlag(int Option, bool Persist = false)
        {
            return Application.GetOptionFlag(Option, Persist);
        }

        public FileDialog get_FileDialog(MsoFileDialogType Type)
        {
            return Application.get_FileDialog(Type);
        }

        public void Help(string HelpFile = "vbapp10.chm", int ContextID = 0)
        {
            Application.Help(HelpFile, ContextID);
        }

        public void LaunchPublishSlidesDialog(string SlideLibraryUrl)
        {
            Application.LaunchPublishSlidesDialog(SlideLibraryUrl);
        }

        public void LaunchSendToPPTDialog(ref object SlideUrls)
        {
            Application.LaunchSendToPPTDialog(SlideUrls);
        }

        public void LaunchSpelling(ppt.DocumentWindow pWindow)
        {
            Application.LaunchSpelling(pWindow);
        }

        public ppt.Theme OpenThemeFile(string themeFileName)
        {
            return Application.OpenThemeFile(themeFileName);
        }

        public dynamic PPFileDialog(ppt.PpFileDialogType Type)
        {
            return Application.PPFileDialog(Type);
        }

        public void Quit()
        {
            Application.Quit();
        }

        public dynamic Run(string MacroName, params object[] safeArrayOfParams)
        {
            return Application.Run(MacroName, safeArrayOfParams);
        }

        public void SetOptionFlag(int Option, bool State, bool Persist = false)
        {
            Application.SetOptionFlag(Option, State, Persist);
        }

        public void SetPerfMarker(int Marker)
        {
            Application.SetPerfMarker(Marker);
        }

        public void StartNewUndoEntry()
        {
            Application.StartNewUndoEntry();
        }

    }
}
