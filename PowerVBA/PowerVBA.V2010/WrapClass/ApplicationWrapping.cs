using VBA = Microsoft.Vbe.Interop;
using PowerVBA.Core.Interface;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Collections;
using PowerVBA.Core.Wrap;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Core.Connector;
using System;

namespace PowerVBA.V2010.WrapClass

{
    [Wrapped(typeof(Application))]
    public class ApplicationWrapping : ApplicationWrappingBase
    {


        public Application Application { get; }
        public ApplicationWrapping(Application application)
        {
            this.Application = application;
        }

        public override PPTVersion ClassVersion => PPTVersion.PPT2010;

        public void Help(string HelpFile = "vbapp10.chm", int ContextID = 0)
        {
            Application.Help(HelpFile, ContextID);
        }

        public void Quit()
        {
            Application.Quit();
        }

        public dynamic Run(string MacroName, params object[] safeArrayOfParams)
        {
            return Application.Run(MacroName, safeArrayOfParams);
        }

        public dynamic PPFileDialog(PpFileDialogType Type)
        {
            return Application.PPFileDialog(Type);
        }

        public void LaunchSpelling(DocumentWindow pWindow)
        {
            Application.LaunchSpelling(pWindow);
        }

        public void Activate()
        {
            Application.Activate();
        }

        public bool GetOptionFlag(int Option, bool Persist = false)
        {
            return Application.GetOptionFlag(Option, Persist);
        }

        public void SetOptionFlag(int Option, bool State, bool Persist = false)
        {
            Application.SetOptionFlag(Option, State, Persist);
        }

        public FileDialog get_FileDialog(MsoFileDialogType Type)
        {
            return Application.get_FileDialog(Type);
        }

        public void SetPerfMarker(int Marker)
        {
            Application.SetPerfMarker(Marker);
        }

        public void LaunchPublishSlidesDialog(string SlideLibraryUrl)
        {
            Application.LaunchPublishSlidesDialog(SlideLibraryUrl);
        }

        public void LaunchSendToPPTDialog(ref object SlideUrls)
        {
            Application.LaunchSendToPPTDialog(ref SlideUrls);
        }

        public void StartNewUndoEntry()
        {
            Application.StartNewUndoEntry();
        }

        public Presentations Presentations => Application.Presentations;
        public DocumentWindows Windows => Application.Windows;
        public dynamic Dialogs => Application.Dialogs;
        public DocumentWindow ActiveWindow => Application.ActiveWindow;
        public Presentation ActivePresentation => Application.ActivePresentation;
        public SlideShowWindows SlideShowWindows => Application.SlideShowWindows;
        public CommandBars CommandBars => Application.CommandBars;
        public string Path => Application.Path;
        public string Name => Application.Name;
        public string Caption { set { Application.Caption = value; } get { return Application.Caption; } }
        public Assistant Assistant => Application.Assistant;
        public FileSearch FileSearch => Application.FileSearch;
        public IFind FileFind => Application.FileFind;
        public string Build => Application.Build;
        public string Version => Application.Version;
        public string OperatingSystem => Application.OperatingSystem;
        public string ActivePrinter => Application.ActivePrinter;
        public int Creator => Application.Creator;
        public AddIns AddIns => Application.AddIns;
        public VBA.VBE VBE => Application.VBE;
        public float Left { set { Application.Left = value; } get { return Application.Left; } }
        public float Top { set { Application.Top = value; } get { return Application.Top; } }
        public float Width { set { Application.Width = value; } get { return Application.Width; } }
        public float Height { set { Application.Height = value; } get { return Application.Height; } }
        public PpWindowState WindowState { set { Application.WindowState = value; } get { return Application.WindowState; } }
        public MsoTriState Visible { set { Application.Visible = value; } get { return Application.Visible; } }
        public int HWND => Application.HWND;
        public MsoTriState Active => Application.Active;
        public AnswerWizard AnswerWizard => Application.AnswerWizard;
        public COMAddIns COMAddIns => Application.COMAddIns;
        public string ProductCode => Application.ProductCode;
        public DefaultWebOptions DefaultWebOptions => Application.DefaultWebOptions;
        public LanguageSettings LanguageSettings => Application.LanguageSettings;
        public MsoDebugOptions MsoDebugOptions => Application.MsoDebugOptions;
        public MsoTriState ShowWindowsInTaskbar { set { Application.ShowWindowsInTaskbar = value; } get { return Application.ShowWindowsInTaskbar; } }
        public dynamic Marker => Application.Marker;
        public MsoFeatureInstall FeatureInstall { set { Application.FeatureInstall = value; } get { return Application.FeatureInstall; } }
        public MsoTriState DisplayGridLines { set { Application.DisplayGridLines = value; } get { return Application.DisplayGridLines; } }
        public MsoAutomationSecurity AutomationSecurity { set { Application.AutomationSecurity = value; } get { return Application.AutomationSecurity; } }
        public NewFile NewPresentation => ((_Application)Application).NewPresentation;
        public PpAlertLevel DisplayAlerts { set { Application.DisplayAlerts = value; } get { return Application.DisplayAlerts; } }
        public MsoTriState ShowStartupDialog { set { Application.ShowStartupDialog = value; } get { return Application.ShowStartupDialog; } }
        public AutoCorrect AutoCorrect => Application.AutoCorrect;
        public Options Options => Application.Options;
        public bool DisplayDocumentInformationPanel { set { Application.DisplayDocumentInformationPanel = value; } get { return Application.DisplayDocumentInformationPanel; } }
        public IAssistance Assistance => Application.Assistance;
        public int ActiveEncryptionSession => Application.ActiveEncryptionSession;
        public FileConverters FileConverters => Application.FileConverters;
        public SmartArtLayouts SmartArtLayouts => Application.SmartArtLayouts;
        public SmartArtQuickStyles SmartArtQuickStyles => Application.SmartArtQuickStyles;
        public SmartArtColors SmartArtColors => Application.SmartArtColors;
        public ProtectedViewWindows ProtectedViewWindows => Application.ProtectedViewWindows;
        public ProtectedViewWindow ActiveProtectedViewWindow => Application.ActiveProtectedViewWindow;
        public bool IsSandboxed => Application.IsSandboxed;
        public ResampleMediaTasks ResampleMediaTasks => Application.ResampleMediaTasks;
        public MsoFileValidationMode FileValidation { set { Application.FileValidation = value; } get { return Application.FileValidation; } }
    }
}