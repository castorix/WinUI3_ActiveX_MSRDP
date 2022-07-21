using GlobalStructures;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using MSRDP;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

using static MSRDP.MSRDPTools;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace WinUI3_ActiveX_MSRDP
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public static void SafeRelease<T>(ref T comObject) where T : class
        {
            T t = comObject;
            comObject = default(T);
            if (null != t)
            {
                if (Marshal.IsComObject(t))
                    Marshal.ReleaseComObject(t);
            }
        }

        [DllImport("Atl.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int AtlAxGetHost(IntPtr h, [MarshalAs(UnmanagedType.IUnknown)] out object pp);

        [DllImport("Atl.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int AtlAxCreateControl(string lpszName, IntPtr hWnd, IntPtr pStream, [MarshalAs(UnmanagedType.IUnknown)] out object ppUnkContainer);

        [DllImport("Atl.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool AtlAxWinInit();

        [DllImport("Atl.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int AtlAxGetControl(IntPtr h, [MarshalAs(UnmanagedType.IUnknown)] out object pp);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool TerminateProcess(IntPtr hProcess, uint uExitCode);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr GetCurrentProcess();

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int GetSystemMetrics(int nIndex);

        public const int SM_CXSCREEN = 0;
        public const int SM_CYSCREEN = 1;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool ShowWindow(IntPtr hWnd, int nShowCmd);

        public const int SW_HIDE = 0;
        public const int SW_SHOWNORMAL = 1;
        public const int SW_SHOWMINIMIZED = 2;
        public const int SW_SHOWMAXIMIZED = 3;
        public const int SW_SHOWNOACTIVATE = 4;
        public const int SW_SHOW = 5;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprcUpdate, IntPtr hrgnUpdate, uint flags);

        public const int RDW_INVALIDATE = 0x0001;
        public const int RDW_INTERNALPAINT = 0x0002;
        public const int RDW_ERASE = 0x0004;

        public const int RDW_VALIDATE = 0x0008;
        public const int RDW_NOINTERNALPAINT = 0x0010;
        public const int RDW_NOERASE = 0x0020;

        public const int RDW_NOCHILDREN = 0x0040;
        public const int RDW_ALLCHILDREN = 0x0080;

        public const int RDW_UPDATENOW = 0x0100;
        public const int RDW_ERASENOW = 0x0200;

        public const int RDW_FRAME = 0x0400;
        public const int RDW_NOFRAME = 0x0800;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        public const int SWP_NOSIZE = 0x0001;
        public const int SWP_NOMOVE = 0x0002;
        public const int SWP_NOZORDER = 0x0004;
        public const int SWP_NOREDRAW = 0x0008;
        public const int SWP_NOACTIVATE = 0x0010;
        public const int SWP_FRAMECHANGED = 0x0020;  /* The frame changed: send WM_NCCALCSIZE */
        public const int SWP_SHOWWINDOW = 0x0040;
        public const int SWP_HIDEWINDOW = 0x0080;
        public const int SWP_NOCOPYBITS = 0x0100;
        public const int SWP_NOOWNERZORDER = 0x0200;  /* Don't do owner Z ordering */
        public const int SWP_NOSENDCHANGING = 0x0400;  /* Don't send WM_WINDOWPOSCHANGING */
        public const int SWP_DRAWFRAME = SWP_FRAMECHANGED;
        public const int SWP_NOREPOSITION = SWP_NOOWNERZORDER;
        public const int SWP_DEFERERASE = 0x2000;
        public const int SWP_ASYNCWINDOWPOS = 0x4000;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        public const int GW_HWNDFIRST = 0;
        public const int GW_HWNDLAST = 1;
        public const int GW_HWNDNEXT = 2;
        public const int GW_HWNDPREV = 3;
        public const int GW_OWNER = 4;
        public const int GW_CHILD = 5;
        public const int GW_ENABLEDPOPUP = 6;  

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        public delegate int SUBCLASSPROC(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, uint dwRefData);

        [DllImport("Comctl32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SetWindowSubclass(IntPtr hWnd, SUBCLASSPROC pfnSubclass, uint uIdSubclass, uint dwRefData);

        [DllImport("Comctl32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int DefSubclassProc(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam);

        public const int WM_SIZE = 0x0005;
        public const int WM_PAINT = 0x000F;
        public const int WM_ERASEBKGND = 0x0014;
        public const int WM_NCCALCSIZE = 0x0083;
        public const int WM_SETCURSOR = 0x0020;
        public const int WM_NCHITTEST = 0x0084;
        public const int WM_MOUSEACTIVATE = 0x0021;

        public const int MA_ACTIVATE = 1;
        public const int MA_ACTIVATEANDEAT = 2;
        public const int MA_NOACTIVATE = 3;
        public const int MA_NOACTIVATEANDEAT = 4;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool FillRect(IntPtr hdc, [In] ref RECT rect, IntPtr hbrush);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateSolidBrush(int crColor);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool DeleteObject([In] IntPtr hObject);

        // https://csms.fos.transas.com/api/rdp/get/rdp_test_page.html
        private string _sRDPFile = "RDP.xml";
        RDP _RDPData = null;

        private IntPtr hWnd = IntPtr.Zero;
        IMsRdpClient7 _pRDPClient = null;
        System.Runtime.InteropServices.ComTypes.IConnectionPoint _pConnectionPoint = null;
        int _nCookie = 0;
        bool m_bNotifyFullScreenMode = false;

        public MainWindow()
        {
            this.InitializeComponent();
            hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);

            this.VisibilityChanged += MainWindow_VisibilityChanged;
            this.Closed += MainWindow_Closed;

            LoadData();            
        }
        
        private SUBCLASSPROC SubClassDelegate;

        private bool bInitialized = false;
        private void MainWindow_VisibilityChanged(object sender, WindowVisibilityChangedEventArgs args)
        {
            if (!bInitialized)
            { 
                if (AtlAxWinInit())
                {
                    HRESULT hr = HRESULT.S_OK;
                    ContainerPanel1.ContainerPanelInit("AtlAxWin", "{" + CLSID_MsRdpClient11NotSafeForScripting.ToString() + "}", hWnd);
                    ContainerPanel1.SetOpacity(ContainerPanel1.hWndContainer, 100);
                    ContainerPanel1.Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/Remote_desktop_connection_icon.webp"));
                    ContainerPanel1.Stretch = Stretch.Uniform;
                    ContainerPanel1.Hide(false);
                    object pUnknown = null;
                    AtlAxGetControl(ContainerPanel1.hWndContainer, out pUnknown);
                    if (pUnknown != null)
                    {
                        _pRDPClient = (IMsRdpClient7)pUnknown;
                        if (_pRDPClient != null)
                        {
                            System.Runtime.InteropServices.ComTypes.IConnectionPointContainer pConnectionPointContainer = null;
                            pConnectionPointContainer = (System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)_pRDPClient;
                            Guid guidEvents = typeof(IMsTscAxEvents).GUID;
                            pConnectionPointContainer.FindConnectionPoint(ref guidEvents, out _pConnectionPoint);
                            //Marshal.ReleaseComObject(pConnectionPointContainer);
                            CEventSink pEventSink = new CEventSink(this);
                            _pConnectionPoint.Advise(pEventSink, out _nCookie);
                            //_pConnectionPoint.Advise(Marshal.GetComInterfaceForObject(pEventSink, typeof(IMsTscAxEvents)), out _nCookie);
                            //Marshal.ReleaseComObject(_pConnectionPoint);

                            string sVersion = null;
                            hr = _pRDPClient.get_Version(out sVersion);                          

                            //hr = _pRDPClient.put_Server("54.220.223.143");
                            //hr = _pRDPClient.put_UserName("SimCloudTest");
                            IMsRdpClientAdvancedSettings7 pSettings = null;
                            hr = _pRDPClient.get_AdvancedSettings8(out pSettings);
                            if (pSettings != null)
                            {
                                //hr = pSettings.put_ClearTextPassword("SimTest!PSWD");
                                hr = pSettings.put_EnableCredSspSupport(true);
                                hr = pSettings.put_DisplayConnectionBar(true);
                                hr = pSettings.put_ConnectionBarShowMinimizeButton(false);
                                hr = pSettings.put_AuthenticationLevel(2);
                                hr = pSettings.put_SmartSizing(true);                               
                                //hr = pSettings.put_NegotiateSecurityLayer(false);

                                //hr = _pRDPClient.Connect();

                                IMsRdpClientNonScriptable5 pNS = (IMsRdpClientNonScriptable5)pUnknown;
                                if (pNS != null)
                                {
                                    //hr = pNS.put_PromptForCredentials(false);
                                    //hr = pNS.put_WarnAboutSendingCredentials(false);

                                    //hr = pNS.put_PromptForCredsOnClient(false);
                                    //hr = pNS.put_AllowPromptingForCredentials(false);

                                    // AllowCredentialSaving = false;                                   
                                    // MarkRdpSettingsSecure = true;
                                }
                                Marshal.ReleaseComObject(pSettings);
                            }
                        }
                        SubClassDelegate = new SUBCLASSPROC(WindowSubClass);
                        IntPtr hWndToSubclass = GetWindow(ContainerPanel1.hWndContainer, GW_CHILD);
                        bool bRet = SetWindowSubclass(hWndToSubclass, SubClassDelegate, 0, 0);
                        bInitialized = true;
                    }                                     
                }
            }
        }

        private void MainWindow_Closed(object sender, WindowEventArgs args)
        {
            HRESULT hr = HRESULT.S_OK;
            SaveData();
            if (_pRDPClient != null)
            {
                short isConnected = 0;
                hr = _pRDPClient.get_Connected(out isConnected);
                if (isConnected != 0)
                    hr = _pRDPClient.Disconnect();
                if (_pConnectionPoint != null)
                    _pConnectionPoint.Unadvise(_nCookie);
                SafeRelease(ref _pConnectionPoint);
                SafeRelease(ref _pRDPClient);
                // To be fixed : process hangs...
                TerminateProcess(GetCurrentProcess(), 0);
            }
        }

        private async void btnConnect_Click(object sender, RoutedEventArgs e)
        {
            HRESULT hr = HRESULT.S_OK;
            //bool bFullScreen = false;
            //hr = pRDPClient.get_FullScreen(out bFullScreen);
            bool bFullScreen = tsFullscreen.IsOn ? true : false;
            if (!bFullScreen)
            {
                hr = _pRDPClient.put_FullScreen(false);
                // E_FAIL if called after Connect()
                //hr = pRDPClient.put_DesktopWidth((int)ContainerPanel1.ActualWidth);
                //hr = pRDPClient.put_DesktopHeight((int)ContainerPanel1.ActualHeight);

                // Better with SmartSizing
                int nScreenWidth = GetSystemMetrics(SM_CXSCREEN);
                int nScreenHeight = GetSystemMetrics(SM_CYSCREEN);
                hr = _pRDPClient.put_DesktopWidth(nScreenWidth);
                hr = _pRDPClient.put_DesktopHeight(nScreenHeight);
            }
            else
            {
                hr = _pRDPClient.put_FullScreen(true);
                int nScreenWidth = GetSystemMetrics(SM_CXSCREEN);
                int nScreenHeight = GetSystemMetrics(SM_CYSCREEN);
                hr = _pRDPClient.put_DesktopWidth(nScreenWidth);
                hr = _pRDPClient.put_DesktopHeight(nScreenHeight);
            }
            short isConnected = 0;
            hr = _pRDPClient.get_Connected(out isConnected);
            if (isConnected == 0)
            {
                hr = _pRDPClient.put_Server(tbServer.Text);
                hr = _pRDPClient.put_UserName(tbUser.Text);
                IMsRdpClientAdvancedSettings7 pSettings = null;
                hr = _pRDPClient.get_AdvancedSettings8(out pSettings);
                if (pSettings != null)
                {
                    hr = pSettings.put_ClearTextPassword(tbPassword.Password);                  
                    Marshal.ReleaseComObject(pSettings);
                }

                //IMsRdpClientTransportSettings pTransportSettings = null;
                //hr = _pRDPClient.get_TransportSettings(out pTransportSettings);
                //if (pTransportSettings != null)
                //{
                //    hr = pTransportSettings.put_GatewayHostname(tbGateway.Text);
                //    Marshal.ReleaseComObject(pTransportSettings);
                //}

                hr = _pRDPClient.Connect();
                if (hr != HRESULT.S_OK)
                {
                    string sMessage = string.Format("Connexion error ({0:X})", hr);
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sMessage, "Error");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                    _ = await md.ShowAsync();
                }               
            }
            else
            {
                Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog("RDP client already connected", "Information");
                WinRT.Interop.InitializeWithWindow.Initialize(md, hWnd);
                _ = await md.ShowAsync();
            }
        }

        private void tsFullscreen_Toggled(object sender, RoutedEventArgs e)
        {
            HRESULT hr = HRESULT.S_OK;
            ToggleSwitch ts = sender as ToggleSwitch;

            if (!m_bNotifyFullScreenMode)
            {
                bool bFullScreen = ts.IsOn ? true : false;
                if (bFullScreen)
                {
                    hr = _pRDPClient.put_FullScreen(true);
                    int nScreenWidth = GetSystemMetrics(SM_CXSCREEN);
                    int nScreenHeight = GetSystemMetrics(SM_CYSCREEN);
                    hr = _pRDPClient.put_DesktopWidth(nScreenWidth);
                    hr = _pRDPClient.put_DesktopHeight(nScreenHeight);
                }
                else
                {
                    hr = _pRDPClient.put_FullScreen(false);
                    //hr = pRDPClient.put_DesktopWidth((int)ContainerPanel1.ActualWidth);
                    //hr = pRDPClient.put_DesktopHeight((int)ContainerPanel1.ActualHeight);

                    // Better with SmartSizing
                    int nScreenWidth = GetSystemMetrics(SM_CXSCREEN);
                    int nScreenHeight = GetSystemMetrics(SM_CYSCREEN);
                    hr = _pRDPClient.put_DesktopWidth(nScreenWidth);
                    hr = _pRDPClient.put_DesktopHeight(nScreenHeight);
                }
            }
            else
                m_bNotifyFullScreenMode = false;
        }

        public void OnConnecting()
        {
            btnConnect.IsEnabled = false;
        }

        public void OnConnected()
        {
            btnConnect.IsEnabled = false;
            ContainerPanel1.Show(true);
        }

        public void OnDisconnected()
        {
            btnConnect.IsEnabled = true;
            ContainerPanel1.Source = new Microsoft.UI.Xaml.Media.Imaging.BitmapImage(new Uri("ms-appx:///Assets/Remote_desktop_connection_icon.webp"));
            ContainerPanel1.Hide(false);
        }

        public void OnEnterFullScreenMode()
        {
            if (!tsFullscreen.IsOn)
            {
                m_bNotifyFullScreenMode = true;
                tsFullscreen.IsOn = true;
            }
            ShowWindow(hWnd, SW_SHOWMINIMIZED);
        }

        public void OnLeaveFullScreenMode()
        {
            if (tsFullscreen.IsOn)
            {
                m_bNotifyFullScreenMode = true;
                tsFullscreen.IsOn = false;
            }
            //AtlAxWin
            // ATL:77F61000
            //  UIMainClass
            //   UIContainerClass
            //     OPContainerClass
            //      OPWindowClass
            ShowWindow(hWnd, SW_SHOWNORMAL);
            //RedrawWindow(hWnd, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_UPDATENOW | RDW_ALLCHILDREN);
            RECT rc;
            GetWindowRect(ContainerPanel1.hWndContainer, out rc);
            SetWindowPos(ContainerPanel1.hWndContainer, (IntPtr)(-1), 0, 0, rc.right - rc.left - 1, rc.bottom - rc.top - 1, SWP_NOZORDER | SWP_NOMOVE | SWP_SHOWWINDOW | SWP_FRAMECHANGED);
            SetWindowPos(ContainerPanel1.hWndContainer, (IntPtr)(-1), 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER | SWP_NOMOVE | SWP_SHOWWINDOW | SWP_FRAMECHANGED);
        }

        public void SaveData()
        {
            string sDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string sFile = sDirectory + _sRDPFile;
            WriteObject(sFile);
        }

        public void LoadData()
        {
            string sDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string sFile = sDirectory + _sRDPFile;
            ReadObject(sFile);
        }

        // From MSDN
        [System.Runtime.Serialization.DataContract(Name = "RDP")]
        class RDP : System.Runtime.Serialization.IExtensibleDataObject
        {
            [System.Runtime.Serialization.DataMember()]
            public string Server;
            [System.Runtime.Serialization.DataMember]
            public string User;
            [System.Runtime.Serialization.DataMember()]
            public string Password;

            public RDP(string sServer, string sUser, string sPassword)
            {
                Server = sServer;
                User = sUser;
                Password = sPassword;
            }

            private System.Runtime.Serialization.ExtensionDataObject extensionData_Value;

            public System.Runtime.Serialization.ExtensionDataObject ExtensionData
            {
                get
                {
                    return extensionData_Value;
                }
                set
                {
                    extensionData_Value = value;
                }
            }
        }

        public void WriteObject(string sFileName)
        {
            FileStream writer = new FileStream(sFileName, FileMode.Create);
            System.Runtime.Serialization.DataContractSerializer ser = new System.Runtime.Serialization.DataContractSerializer(typeof(RDP));
            _RDPData.Server = tbServer.Text;
            _RDPData.User = tbUser.Text;
            _RDPData.Password = tbPassword.Password;
            ser.WriteObject(writer, _RDPData);
            writer.Close();
        }

        public void ReadObject(string sFileName)
        {
            try
            {
                FileStream fs = new FileStream(sFileName, FileMode.Open);
                System.Xml.XmlDictionaryReader reader = System.Xml.XmlDictionaryReader.CreateTextReader(fs, new System.Xml.XmlDictionaryReaderQuotas());
                System.Runtime.Serialization.DataContractSerializer ser = new System.Runtime.Serialization.DataContractSerializer(typeof(RDP));
                _RDPData = (RDP)ser.ReadObject(reader, true);
                tbServer.Text = _RDPData.Server;
                tbUser.Text = _RDPData.User;
                tbPassword.Password = _RDPData.Password;
                reader.Close();
                fs.Close();
            }
            catch (Exception ex)
            {
                var sEx = ex.Message;
                _RDPData = new RDP("54.220.223.143", "SimCloudTest", "SimTest!PSWD");
                tbServer.Text = _RDPData.Server;
                tbUser.Text = _RDPData.User;
                tbPassword.Password = _RDPData.Password;
            }
        }

        private int WindowSubClass(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, uint dwRefData)
        {
            switch (uMsg)
            {
                case WM_MOUSEACTIVATE:
                    {
                        //SetFocus(GetWindow(hWnd, GW_OWNER));
                        //return MA_ACTIVATE;
                        RECT rc;
                        GetWindowRect(ContainerPanel1.hWndContainer, out rc);
                        SetWindowPos(ContainerPanel1.hWndContainer, (IntPtr)(-1), 0, 0, rc.right - rc.left - 1, rc.bottom - rc.top - 1, SWP_NOZORDER | SWP_NOMOVE | SWP_SHOWWINDOW | SWP_FRAMECHANGED);
                        SetWindowPos(ContainerPanel1.hWndContainer, (IntPtr)(-1), 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER | SWP_NOMOVE | SWP_SHOWWINDOW | SWP_FRAMECHANGED);
                    }
                    break;
                    
                    //case WM_ERASEBKGND:
                    //    {
                    //        RECT rect;
                    //        GetClientRect(hWnd, out rect);

                    //        //int nRet = ExcludeClipRect(wParam, rect.left, rect.top, rect.right, rect.bottom);

                    //        //IntPtr hBrush = CreateSolidBrush(System.Drawing.ColorTranslator.ToWin32(System.Drawing.Color.Red));
                    //        //IntPtr hBrush = CreateSolidBrush(System.Drawing.ColorTranslator.ToWin32(System.Drawing.Color.FromArgb(255, 32, 32, 32)));
                    //        IntPtr hBrush = CreateSolidBrush(System.Drawing.ColorTranslator.ToWin32(System.Drawing.Color.Blue));
                    //        FillRect(wParam, ref rect, hBrush);
                    //        DeleteObject(hBrush);
                    //        return 1;
                    //    }
                    //    break;
            }
            return DefSubclassProc(hWnd, uMsg, wParam, lParam);
        }
    }
}
