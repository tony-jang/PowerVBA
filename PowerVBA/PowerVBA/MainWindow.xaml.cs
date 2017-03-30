using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerVBA.Windows;
using PowerVBA.Core.Connector;
using Microsoft.Win32;
using System.Windows.Threading;
using PowerVBA.Core.AvalonEdit;
using PowerVBA.Windows.AddWindows;
using static PowerVBA.Global.Globals;
using PowerVBA.V2013.Connector;
using PowerVBA.Core.Extension;
using PowerVBA.Core.Wrap.WrapBase;
using PowerVBA.Wrap;
using PowerVBA.Controls.Customize;
using PowerVBA.Core.Controls;
using System.Threading;
using System.ComponentModel;
using System.Collections.Generic;
using ICSharpCode.AvalonEdit.Document;
using PowerVBA.V2013.Wrap.WrapClass;
using PowerVBA.Controls.Tools;
using PowerVBA.Codes;
using PowerVBA.Codes.TypeSystem;
using System.Diagnostics;
using PowerVBA.Core.AvalonEdit.Parser;
using PowerVBA.Core.Error;
using ICSharpCode.AvalonEdit;

namespace PowerVBA
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : ChromeWindow
    {
        PPTConnectorBase Connector;
        StartupWindow suWindow = new StartupWindow();
        BackgroundWorker bg;
        Thread loadThread;
        SQLiteConnector dbConnector;
        CodeInfo codeInfo;
        public MainWindow()
        {
            InitializeComponent();
            Stopwatch sw = new Stopwatch();

            codeInfo = new CodeInfo();

            sw.Start();
            new VBAParser(codeInfo).Parse(@"Imports FmodExample.FMOD
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Threading

Public Class FmodPlayer
    Public Enum States As Integer
        [Default] = 0
        [Playing] = 1
        [Stop] = 2
        [Paused] = 3
    End Enum

    Private updateThread As Thread

    Private _Fmod As New FMOD.System
    Private _Sound As New Sound
    Private _Channel As New Channel

    Private _Looping As Boolean = False
    Private _State As States = States.Default
    Private _Length As UInteger = 0
    Private _Position As UInteger = 0
    Private _Volume As Single = 0.0F
    Private _Pan As Single = 0.0F

    Public Effect As SoundEffect
    Public Equalizer As SoundEqualizer

    Sub New()
        Factory.System_Create(_Fmod)
        _Fmod.init(32, INITFLAGS.NORMAL, IntPtr.Zero)
        _Fmod.setStreamBufferSize(6400 * 1024, FMOD.TIMEUNIT.RAWBYTES)

        Effect = New SoundEffect(_Fmod)
        Equalizer = New SoundEqualizer(_Fmod, _Channel)

        updateThread = New Thread(New ThreadStart(Sub() UpdateWorker()))
        updateThread.Start()
    End Sub

    Public Sub Dispose()
        _Fmod.close()

        _Fmod = Nothing
    End Sub

    Private Sub UpdateWorker()
        Dim AREvt As New AutoResetEvent(False)

        While (Not _Fmod Is Nothing)
            If _Looping And _State <> States.Default Then
                If Position = Length Then
                    Dim boPaused As Boolean = False
                    boPaused = (_State = States.Paused)
                    _Fmod.playSound(CHANNELINDEX.FREE Or CHANNELINDEX.REUSE, _Sound, boPaused, _Channel)
                    Volume = Volume
                End If
            End If
            AREvt.WaitOne(10, True)
        End While
    End Sub

#Region ""프로퍼티""
    Public ReadOnly Property State As States
        Get
            Return _State
        End Get
    End Property

    Public ReadOnly Property Length As UInteger
        Get
            Return _Length
        End Get
    End Property

    Public Property Position As UInteger
        Get
            _Channel.getPosition(_Position, TIMEUNIT.MS)
            Return _Position
        End Get
        Set(value As UInteger)
            If _State <> States.Default Then
                _Position = value
                If _Position > _Length Then _Position = _Length

                _Channel.setPosition(_Position, TIMEUNIT.MS)
            End If
        End Set
    End Property

    Public Property Volume As Single
        Get
            Return _Volume
        End Get
        Set(value As Single)
            _Channel.setVolume(value)
            _Channel.getVolume(_Volume)
        End Set
    End Property

    Public Property Pan As Single
        Get
            Return _Pan
        End Get
        Set(value As Single)
            _Channel.setPan(value)
            _Channel.getPan(_Pan)
        End Set
    End Property

    Public ReadOnly Property Spectrum(SpectrumSize As Integer) As Single()
        Get
            Dim _Spec(SpectrumSize) As Single
            Dim _Spec1(SpectrumSize) As Single
            Dim _Spec2(SpectrumSize) As Single
            _Channel.getSpectrum(_Spec1, SpectrumSize, 0, DSP_FFT_WINDOW.RECT)
            _Channel.getSpectrum(_Spec2, SpectrumSize, 1, DSP_FFT_WINDOW.RECT)

            For i As Integer = 0 To SpectrumSize - 1
                _Spec(i) = ((_Spec1(i) + _Spec2(i)) / 2)
            Next

            Return _Spec
        End Get
    End Property
#End Region

    Public Function PlaySound(fileName As String, boPaused As Boolean, looping As Boolean) As Boolean
        _Sound.release()
        Dim result As RESULT = _Fmod.createSound(fileName, MODE.DEFAULT Or MODE.HARDWARE, _Sound)

        If result = FMOD.RESULT.OK Then
            Dim result2 As RESULT = _Fmod.playSound(CHANNELINDEX.FREE Or CHANNELINDEX.REUSE, _Sound, boPaused, _Channel)
            If result2 = FMOD.RESULT.OK Then
                _Sound.getLength(_Length, TIMEUNIT.MS)
                _Channel.getVolume(_Volume)
                _Channel.getPan(_Pan)
                _Looping = looping
                _State = States.Playing
                Return True
            End If
        End If

        Return False
    End Function

    Public Function[Stop]() As Boolean
        If _State <> States.Default Then
            Dim result As RESULT = _Channel.stop

            If result = FMOD.RESULT.OK Then
                _State = States.Stop
                Return True
            End If
        End If

        Return False
    End Function

    Public Function[Pause]() As Boolean
        If _State = States.Playing Then
            Dim result As RESULT = _Channel.setPaused(True)

            If result = FMOD.RESULT.OK Then
                _State = States.Paused
                Return True
            End If
        End If

        Return False
    End Function

    Public Function[Resume]() As Boolean
        If _State = States.Paused Then
            Dim result As RESULT = _Channel.setPaused(False)

            If result = FMOD.RESULT.OK Then
                _State = States.Playing
                Return True
            End If
        End If

        Return False
    End Function

    Public Class SoundEqualizer
        Private _fmod As FMOD.System
        Private _Channel As Channel

        Private eqa1 As FMOD.DSPConnection
        Private eqa2 As FMOD.DSPConnection
        Private eqa3 As FMOD.DSPConnection
        Private eqa4 As FMOD.DSPConnection
        Private eqa5 As FMOD.DSPConnection
        Private eqa6 As FMOD.DSPConnection
        Private eqa7 As FMOD.DSPConnection
        Private eqa8 As FMOD.DSPConnection
        Private eqa9 As FMOD.DSPConnection
        Private eqa10 As FMOD.DSPConnection

        Private Eq0 As FMOD.DSP
        Private Eq1 As FMOD.DSP
        Private Eq2 As FMOD.DSP
        Private Eq3 As FMOD.DSP
        Private Eq4 As FMOD.DSP
        Private Eq5 As FMOD.DSP
        Private Eq6 As FMOD.DSP
        Private Eq7 As FMOD.DSP
        Private Eq8 As FMOD.DSP
        Private Eq9 As FMOD.DSP

        Sub New(system As FMOD.System, channel As Channel)
            Me._fmod = system
            Me._Channel = channel

            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq0)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq1)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq2)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq3)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq4)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq5)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq6)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq7)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq8)
            _fmod.createDSPByType(FMOD.DSP_TYPE.PARAMEQ, Eq9)

            _fmod.addDSP(Eq0, eqa1)
            _fmod.addDSP(Eq1, eqa2)
            _fmod.addDSP(Eq2, eqa3)
            _fmod.addDSP(Eq3, eqa4)
            _fmod.addDSP(Eq4, eqa5)
            _fmod.addDSP(Eq5, eqa6)
            _fmod.addDSP(Eq6, eqa7)
            _fmod.addDSP(Eq7, eqa8)
            _fmod.addDSP(Eq8, eqa9)
            _fmod.addDSP(Eq9, eqa10)

            Eq0.setParameter(FMOD.DSP_PARAMEQ.CENTER, 80)
            Eq0.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq0.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq1.setParameter(FMOD.DSP_PARAMEQ.CENTER, 170)
            Eq1.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq1.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq2.setParameter(FMOD.DSP_PARAMEQ.CENTER, 310)
            Eq2.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 1)
            Eq2.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq3.setParameter(FMOD.DSP_PARAMEQ.CENTER, 600)
            Eq3.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq3.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq4.setParameter(FMOD.DSP_PARAMEQ.CENTER, 1000)
            Eq4.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq4.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq5.setParameter(FMOD.DSP_PARAMEQ.CENTER, 3000)
            Eq5.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq5.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq6.setParameter(FMOD.DSP_PARAMEQ.CENTER, 6000)
            Eq6.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq6.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq7.setParameter(FMOD.DSP_PARAMEQ.CENTER, 12000)
            Eq7.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq7.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq8.setParameter(FMOD.DSP_PARAMEQ.CENTER, 14000)
            Eq8.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq8.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)

            Eq9.setParameter(FMOD.DSP_PARAMEQ.CENTER, 16000)
            Eq9.setParameter(FMOD.DSP_PARAMEQ.BANDWIDTH, 2)
            Eq9.setParameter(FMOD.DSP_PARAMEQ.GAIN, 1)
        End Sub

        Private Sub SetEQ(eq As DSP, value As Single)
            eq.setParameter(DSP_PARAMEQ.GAIN, value)
        End Sub

        Private ReadOnly Property GetEQ(eq As DSP) As Single
            Get
                Dim v As Single
                eq.getParameter(DSP_PARAMEQ.GAIN, v, Nothing, Nothing)

                Return v
            End Get
        End Property

        Public Property Hz32() As Single
            Get
                Return GetEQ(Eq0)
            End Get
            Set(value As Single)
                SetEQ(Eq0, value)
            End Set
        End Property

        Public Property Hz64() As Single
            Get
                Return GetEQ(Eq1)
            End Get
            Set(value As Single)
                SetEQ(Eq1, value)
            End Set
        End Property

        Public Property Hz125() As Single
            Get
                Return GetEQ(Eq2)
            End Get
            Set(value As Single)
                SetEQ(Eq2, value)
            End Set
        End Property

        Public Property Hz250() As Single
            Get
                Return GetEQ(Eq3)
            End Get
            Set(value As Single)
                SetEQ(Eq3, value)
            End Set
        End Property

        Public Property Hz500() As Single
            Get
                Return GetEQ(Eq4)
            End Get
            Set(value As Single)
                SetEQ(Eq4, value)
            End Set
        End Property

        Public Property K1() As Single
            Get
                Return GetEQ(Eq5)
            End Get
            Set(value As Single)
                SetEQ(Eq5, value)
            End Set
        End Property

        Public Property K2() As Single
            Get
                Return GetEQ(Eq6)
            End Get
            Set(value As Single)
                SetEQ(Eq6, value)
            End Set
        End Property

        Public Property K4() As Single
            Get
                Return GetEQ(Eq7)
            End Get
            Set(value As Single)
                SetEQ(Eq7, value)
            End Set
        End Property

        Public Property K8() As Single
            Get
                Return GetEQ(Eq8)
            End Get
            Set(value As Single)
                SetEQ(Eq8, value)
            End Set
        End Property

        Public Property K16() As Single
            Get
                Return GetEQ(Eq9)
            End Get
            Set(value As Single)
                SetEQ(Eq9, value)
            End Set
        End Property

        Public Property Frequency As Long
            Get
                Dim f As Single
                _Channel.getFrequency(f)
                Return f
            End Get
            Set(value As Long)
                _Channel.setFrequency(value)
                Debug.Print(Frequency)
            End Set
        End Property
    End Class

    Public Class SoundEffect
        Private _fmod As FMOD.System

        Private dspconnectiontemp As New DSPConnection
        Private _lowpassDSP As New DSP
        Private _distortionDSP As New DSP
        Private _highpassDSP As New DSP
        Private _chorusDSP As New DSP
        Private _echoDSP As New DSP
        Private _parameqDSP As New DSP
        Private _flangeDSP As New DSP
        Private _pitchshiftDSP As New DSP
        Private _sfxreverbDSP As New DSP

        Sub New(system As FMOD.System)
            Me._fmod = system


            _fmod.createDSPByType(DSP_TYPE.LOWPASS, _lowpassDSP)
            _fmod.createDSPByType(DSP_TYPE.DISTORTION, _distortionDSP)
            _fmod.createDSPByType(DSP_TYPE.HIGHPASS, _highpassDSP)
            _fmod.createDSPByType(DSP_TYPE.CHORUS, _chorusDSP)
            _fmod.createDSPByType(DSP_TYPE.ECHO, _echoDSP)
            _fmod.createDSPByType(DSP_TYPE.PARAMEQ, _parameqDSP)
            _fmod.createDSPByType(DSP_TYPE.FLANGE, _flangeDSP)
            _fmod.createDSPByType(DSP_TYPE.PITCHSHIFT, _pitchshiftDSP)
            _fmod.createDSPByType(DSP_TYPE.SFXREVERB, _sfxreverbDSP)
        End Sub

        Private ReadOnly Property IsDspActive(dsp As DSP) As Boolean
            Get
                Dim active As Boolean = False
                dsp.getActive(active)
                Return active
            End Get
        End Property

        Private Sub SetEffect(_dsp As DSP, active As Boolean)
            If active Then
                If Not IsDspActive(_dsp) Then
                    _fmod.addDSP(_dsp, dspconnectiontemp)
                End If
            Else
                If IsDspActive(_dsp) Then
                    _dsp.remove()
                End If
            End If
        End Sub

        Public Property Lowpass As Boolean
            Get
                Return IsDspActive(_lowpassDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_lowpassDSP, value)
            End Set
        End Property

        Public Property Highpass As Boolean
            Get
                Return IsDspActive(_highpassDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_highpassDSP, value)
            End Set
        End Property

        Public Property Chorus As Boolean
            Get
                Return IsDspActive(_chorusDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_chorusDSP, value)
            End Set
        End Property

        Public Property Echo As Boolean
            Get
                Return IsDspActive(_echoDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_echoDSP, value)
            End Set
        End Property

        Public Property Parameq As Boolean
            Get
                Return IsDspActive(_parameqDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_parameqDSP, value)
            End Set
        End Property

        Public Property Flange As Boolean
            Get
                Return IsDspActive(_flangeDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_flangeDSP, value)
            End Set
        End Property

        Public Property Pitchshift As Boolean
            Get
                Return IsDspActive(_pitchshiftDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_pitchshiftDSP, value)
            End Set
        End Property

        Public Property Sfxreverb As Boolean
            Get
                Return IsDspActive(_sfxreverbDSP)
            End Get
            Set(value As Boolean)
                SetEffect(_sfxreverbDSP, value)
            End Set
        End Property
    End Class
End Class

");


            codeEditor.TextChanged += ParsingCode;

            //foreach (var itm in d)
            //{
            //    MessageBox.Show(itm.ToString());
            //}


            string str = sw.ElapsedMilliseconds + Environment.NewLine + string.Join("\r\n",
                codeInfo.ErrorList.Select((i) => i.Message + " :: " + i.Region.BeginLine));
            MessageBox.Show(str);
            
            bg = new BackgroundWorker();
            bg.DoWork += bg_DoWork;
            bg.RunWorkerCompleted += bg_RunWorkerCompleted;

            bg.WorkerReportsProgress = true;

            this.Closing += MainWindow_Closing;
            this.Loaded += MainWindow_Loaded;

            MenuTabControl.SelectionChanged += MenuTabControl_SelectionChanged;
            CodeTabControl.SelectionChanged += CodeTabControl_SelectionChanged;

            solutionExplorer.OpenRequest += SolutionExplorer_OpenRequest;
            solutionExplorer.OpenPropertyRequest += SolutionExplorer_OpenPropertyRequest;
            

            BackBtn.Click += BackBtn_Click;

            
        }


        #region [  에디터 메뉴 탭  ]

        #region [  홈 탭 이벤트  ]


        #region [  클립보드  ]
        private void BtnCopy_SimpleButtonClicked()
        {
            Clipboard.Clear();
            Clipboard.SetText(((CodeEditor)CodeTabControl.SelectedItem).SelectedText);
        }
        private void BtnPaste_SimpleButtonClicked()
        {
            if (Clipboard.ContainsText())
            {
                string t = Clipboard.GetText();
                CodeEditor editor = ((CodeEditor)CodeTabControl.SelectedItem);

                if (editor.SelectionLength != 0)
                {
                    editor.SelectedText = t;
                }
                else
                {
                    editor.TextArea.Document.Insert(editor.CaretOffset, t);
                }
                
            }            
        }
        #endregion
        
        #region [  작업  ]
        private void BtnUndo_SimpleButtonClicked()
        {
            CodeEditor editor = ((CodeEditor)CodeTabControl.SelectedItem);
            if (editor == null) return;
            if (editor.CanUndo) editor.Undo();
            BtnUndo.IsEnabled = editor.CanUndo;
            BtnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        private void BtnRedo_SimpleButtonClicked()
        {
            CodeEditor editor = ((CodeEditor)CodeTabControl.SelectedItem);
            if (editor == null) return;
            if (editor.CanRedo) editor.Redo();
            BtnUndo.IsEnabled = editor.CanUndo;
            BtnRedo.IsEnabled = editor.CanRedo;
            editor.Focus();
        }

        #endregion
        
        #region [  슬라이드 관리  ]
        private void BtnNewSlide_SimpleButtonClicked()
        {
            int SlideNumber = 0;


            switch (Connector.Version)
            {
                case PPTVersion.PPT2010:

                    break;
                case PPTVersion.PPT2016:
                case PPTVersion.PPT2013:
                    PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

                    if (conn2013.Presentation.Slides.Count != 0) SlideNumber = conn2013.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                    conn2013.Presentation.Slides.AddSlide(SlideNumber + 1, conn2013.Presentation.SlideMaster.CustomLayouts[1]);
                    conn2013.Presentation.Application.ActiveWindow.View.GotoSlide(SlideNumber + 1);
                    break;
                
            }

        }

        private void BtnDelSlide_SimpleButtonClicked()
        {
            PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

            int SlideNumber = conn2013.Presentation.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            if (MessageBox.Show(SlideNumber + "슬라이드를 삭제합니다. 계속하시려면 예로 계속하세요.", "슬라이드 삭제 확인",
                MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                conn2013.Presentation.Slides[SlideNumber].Delete();
            }
        }

        #endregion


        #endregion


        #region [  삽입 탭 이벤트  ]
        private void BtnAddClass_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, AddFileWindow.AddType.Class);

            filewindow.ShowDialog();
        }

        private void BtnAddModule_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, AddFileWindow.AddType.Module);

            filewindow.ShowDialog();
        }

        private void BtnAddForm_SimpleButtonClicked()
        {
            AddFileWindow filewindow = new AddFileWindow(Connector, AddFileWindow.AddType.Form);

            filewindow.ShowDialog();
        }


        private void BtnAddSub_SimpleButtonClicked()
        {
            AddProcedureWindow procWindow = new AddProcedureWindow();

            procWindow.ShowDialog();
        }

        private void BtnAddFunc_SimpleButtonClicked()
        {

        }

        private void BtnAddMouseOverTrigger_SimpleButtonClicked()
        {

        }

        private void BtnAddMouseClickTrigger_SimpleButtonClicked()
        {

        }

        private void BtnAddEnum_SimpleButtonClicked()
        {

        }

        private void BtnAddType_SimpleButtonClicked()
        {

        }
        #endregion

        #endregion


        #region [  윈도우 코드  ]

        public void SetMessage(string Message, int Delay = 3000)
        {
            Thread thr = new Thread(() =>
            {
                Dispatcher.Invoke(() => { tbMessage.Text = Message; });

                Thread.Sleep(Delay);

                Dispatcher.Invoke(() => { tbMessage.Text = "준비"; });
            });

            thr.Start();
        }


        private void SolutionExplorer_OpenPropertyRequest()
        {
            var tabItems = CodeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == "솔루션 탐색기").ToList();
            if (tabItems.Count >= 1)
            {
                CodeTabControl.SelectedItem = tabItems.First();
            }
            else
            {
                CloseableTabItem itm = new CloseableTabItem()
                {
                    Header = "솔루션 탐색기",
                    Content = new ProjectProperty()
                };
                CodeTabControl.Items.Add(itm);
                CodeTabControl.SelectedItem = itm;
            }
        }

        private void SolutionExplorer_OpenRequest(object sender, VBComponentWrappingBase Data)
        {
            var itm = Data.ToVBComponent2013();
            //string value = itm.CodeModule.get_Lines(1, itm.CodeModule.CountOfLines);
            AddCodeTab(Data);

        }

        bool LoadComplete = false;
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Opacity = 0.0;
            loadThread = new Thread(() =>
            {
                suWindow = new StartupWindow();

                Dispatcher.FromThread(loadThread).Invoke(new Action(() =>
                {
                    suWindow.Show();
                }), DispatcherPriority.Background);

                while (!LoadComplete) { }
                suWindow.Close();
            });

            loadThread.SetApartmentState(ApartmentState.STA);
            Dispatcher.Invoke(new Action(() =>
            {
                loadThread.Start();
            }), DispatcherPriority.Background);


            bg.RunWorkerAsync();
        }

        private void bg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Opacity = 1.0;
            LoadComplete = true;

            this.Activate();
        }

        private void bg_DoWork(object sender, DoWorkEventArgs e)
        {
            MainDispatcher = Dispatcher;

            Dispatcher.Invoke(new Action(() =>
            {
                dbConnector = new SQLiteConnector();
                dbConnector.RecentFile_Get().ForEach((fl) => { LVrecentFile.Items.Add(new RecentFileListViewItem(fl)); });
                
                RunVersion.Text = VersionSelector.GetPPTVersion().GetDescription();
            }));
        }

        public void AddCodeTab(string Name)
        {
            CodeEditor codeEditor = new CodeEditor();

            CloseableTabItem codeTab = new CloseableTabItem()
            {
                Header = Name,
                Content = codeEditor
            };

            codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Changed = !(((UndoStack)sender).IsOriginalFile); };
            codeEditor.SaveRequest += () => { SetMessage("저장되었습니다."); };
            CodeTabControl.Items.Add(codeTab);
        }

        private void ParsingCode(object sender, EventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();


            new VBAParser(codeInfo).Parse(((CodeEditor)sender).Text);

            errorList.SetError(codeInfo.ErrorList);
            this.Title = ((float)sw.ElapsedTicks / 1000).ToString();
        }

        public void AddCodeTab(VBComponentWrappingBase component)
        {
            CodeEditor codeEditor = null;
            

            switch (component.ClassVersion)
            {
                case PPTVersion.PPT2010:
                    // TODO : PPT 2010 Version 추가
                    return;
                case PPTVersion.PPT2016:
                case PPTVersion.PPT2013:
                    var comp2013 = component.ToVBComponent2013();

                    var tabItems = CodeTabControl.Items.Cast<CloseableTabItem>().Where((i) => i.Header.ToString() == comp2013.CodeModule.Name).ToList();

                    if (tabItems.Count >= 1)
                    {
                        CodeTabControl.SelectedItem = tabItems.First();
                        return;
                    }

                    codeEditor = new CodeEditor(component);

                    try
                    {
                        var module = comp2013.VBComponent.CodeModule;

                        if (comp2013.VBComponent.CodeModule.CountOfLines == 0) codeEditor.Text = "";
                        else codeEditor.Text = comp2013.VBComponent.CodeModule.get_Lines(1, comp2013.VBComponent.CodeModule.CountOfLines);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("예외가 발생했습니다!");
                    }
                    
                    CloseableTabItem codeTab = new CloseableTabItem()
                    {
                        Header = comp2013.Name,
                        Content = codeEditor
                    };
                    CodeTabControl.Items.Add(codeTab);

                    CodeTabControl.SelectedItem = codeTab;

                    codeEditor.TextChanged += ParsingCode;
                    codeEditor.Document.UndoStack.PropertyChanged += (sender, e) => { codeTab.Changed = !(((UndoStack)sender).IsOriginalFile); };
                    codeEditor.SaveRequest += () => 
                    {
                        SetMessage("저장되었습니다.");
                        comp2013.CodeModule.DeleteLines(1, comp2013.CodeModule.CountOfLines);
                        comp2013.CodeModule.AddFromString(codeEditor.Text);
                    };

                    break;
                
            }
        }

        public void SetName(string customName = "")
        {
            if (customName!= "")
            {
                this.Title = $"{customName} - PowerVBA";
            }
            else if (Connector != null)
            {
                PPTConnector2013 conn2013 = (PPTConnector2013)Connector;

                this.Title = conn2013.Presentation.Name + " - PowerVBA";
                if (conn2013.Presentation.ReadOnly == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    this.Title += " [읽기 전용]";
                }
            }
        }

        private void PPTCloseDetect()
        {
            var result = MessageBox.Show("프레젠테이션이 PowerVBA의 코드가 저장되지 않은 상태에서 닫혔습니다.\r\n코드는 저장한 상태에서 다시 여시겠습니까?\r\n이대로 닫게 되면 코드는 저장되지 않습니다.", "PowerVBA", MessageBoxButton.YesNo);

            if (result == MessageBoxResult.Yes)
            {

            }
            else
            {
                Environment.Exit(0);
            }
        }

        private void BackBtn_Click(object sender, RoutedEventArgs e)
        {
            ProgramTabControl.SelectedIndex = 0;
            this.NoTitle = false;
        }

        private void ProjectFileChange()
        {
            solutionExplorer.Update(Connector);
        }


        private void CodeTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            CloseableTabItem tabItm = ((CloseableTabItem)CodeTabControl.SelectedItem);
            if (tabItm == null) return;
            if (tabItm.Content.GetType() == typeof(CodeEditor))
            {
                var editor = (CodeEditor)tabItm.Content;
                BtnUndo.IsEnabled = editor.CanUndo;
                BtnRedo.IsEnabled = editor.CanRedo;
            }
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            Connector?.Dispose();
            Environment.Exit(0);
        }
        private void MenuTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MenuTabControl.SelectedIndex == 0)
            {
                ProgramTabControl.SelectedIndex = 2;
                MenuTabControl.SelectedIndex = 1;
                this.NoTitle = true;
            }    
        }
        #endregion

        

        #region [  초기 화면  ]
        private void GridOpenAnotherPpt_MouseLeave(object sender, MouseEventArgs e)
        {
            TBOpenAnotherPpt.FontStyle = FontStyles.Normal;
            while (TBOpenAnotherPpt.TextDecorations.Count != 0)
            {
                TBOpenAnotherPpt.TextDecorations.RemoveAt(0);
            }
        }

        private void GridOpenAnotherPpt_MouseMove(object sender, MouseEventArgs e)
        {
            TBOpenAnotherPpt.TextDecorations.Add(TextDecorations.Underline);
        }

        private void GridOpenAnotherPpt_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "프레젠테이션|*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm"
            };
            if (ofd.ShowDialog().Value)
            {
                tbProcessInfoTB.Text = "선택한 템플릿을 적용한 프레젠테이션 프로젝트를 만들고 있습니다.";

                InitalizeConnector(ofd.FileName);
            }
            
        }

        private void BtnNewPPT_Click(object sender, RoutedEventArgs e)
        {

            PPTVersion ver = VersionSelector.GetPPTVersion();

            if (ver != PPTVersion.PPT2013 && ver != PPTVersion.PPT2010 && ver != PPTVersion.PPT2016)
            {
                MessageBox.Show($"죄송합니다. {ver.ToString()}는 지원하지 않는 버전입니다.");
                return;
            }

            tbProcessInfoTB.Text = "선택한 템플릿을 적용한 프레젠테이션 프로젝트를 만들고 있습니다.";

            InitalizeConnector();

        }

        public void InitalizeConnector(string FileLocation = "")
        {
            Dispatcher.Invoke(new Action(() =>
            {
                this.NoTitle = false;
                PPTConnectorBase tmpConn;
                if (FileLocation == "")
                {
                    tmpConn = new PPTConnector2013();
                }
                else
                {
                    tmpConn = new PPTConnector2013(FileLocation);
                }

                tmpConn.PresentationClosed += PPTCloseDetect;
                tmpConn.VBAComponentChange += ProjectFileChange;
                tmpConn.SlideChanged += SlideChangedDetect;
                tmpConn.ShapeChanged += ShapeChangedDetect;
                tmpConn.SectionChanged += SectionChangedDetect;


                Connector = tmpConn;

                PropGrid.SelectedObject = Connector.ToConnector2013().Presentation;
                
                //== Debug ==
                //VBComponentWrappingBase comp;

                //tmpConn.AddClass("Test", out comp);

                //VBComponentWrapping compWrap = (VBComponentWrapping)comp;
                

                //compWrap.CodeModule.InsertLines(1, "Dim A As String\r\nDim B as String\r\nPublic Sub Test\r\n\t\r\nEnd Sub");


                //AddCodeTab(comp);

                //===========
                
                ProgramTabControl.SelectedIndex = 0;
                SetName();
            }), DispatcherPriority.Background);
        }

        private void SectionChangedDetect()
        {
            //projAnalyzer.SectionCount = Connector.ToConnector2013().Presentation.SectionCount;
        }

        private void ShapeChangedDetect()
        {
            int count = 0;
            Connector.ToConnector2013().Presentation.Slides.Cast<Microsoft.Office.Interop.PowerPoint.Slide>().ToList().ForEach((i) => count += i.Shapes.Count);
            projAnalyzer.ShapeCount = count;
        }

        private void SlideChangedDetect()
        {
            projAnalyzer.SlideCount = Connector.ToConnector2013().Presentation.Slides.Count;
        }

        private void BtnNewAssistPPT_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnNewVirtualPPT_Click(object sender, RoutedEventArgs e)
        {
            this.NoTitle = false;
            ProgramTabControl.SelectedIndex = 0;
            SetName("가상 프레젠테이션 1");
        }


        #endregion

        private void DebugBtn_SimpleButtonClicked()
        {
            //CodeTabEditor
            var itm = ((CodeTabEditor)((CloseableTabItem)CodeTabControl.SelectedItem).Content).CodeEditor;
            itm.DeleteIndent();
        }
    }
}
