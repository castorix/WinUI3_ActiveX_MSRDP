using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using GlobalStructures;
using VARIANT_BOOL = System.Boolean;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.CompilerServices;
using WinUI3_ActiveX_MSRDP;

namespace MSRDP
{
    [ComImport]
    [Guid("327bb5cd-834e-4400-aef2-b30e15e5d682")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsTscAx_Redist
    {
        #region <IDispatch>
        int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion       
    }

    [ComImport]
    [Guid("8c11efae-92c3-11d1-bc1e-00c04fa31489")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsTscAx : IMsTscAx_Redist
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        HRESULT put_Server(string pServer);
        HRESULT get_Server(out string pServer);
        [PreserveSig]
        HRESULT put_Domain(string pDomain);
        HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        HRESULT put_UserName(string pUserName);
        HRESULT get_UserName(out string pUserName);
        HRESULT put_DisconnectedText(string pDisconnectedText);
        HRESULT get_DisconnectedText(out string pDisconnectedText);
        HRESULT put_ConnectingText(string pConnectingText);
        HRESULT get_ConnectingText(out string pConnectingText);
        HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        HRESULT put_DesktopWidth(int pVal);
        HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        HRESULT put_DesktopHeight(int pVal);
        HRESULT get_DesktopHeight(out int pVal);
        HRESULT put_StartConnected(int pfStartConnected);
        HRESULT get_StartConnected(out int pfStartConnected);
        HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        HRESULT put_FullScreenTitle(string _arg1);
        HRESULT get_CipherStrength(out int pCipherStrength);
        HRESULT get_Version(out string pVersion);
        HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        HRESULT Connect();
        HRESULT Disconnect();
        HRESULT CreateChannels(string newVal);
        HRESULT SendOnChannel(string chanName, string ChanData);
    }

    [ComImport]
    [Guid("c9d65442-a0f9-45b2-8f73-d61d2db8cbb6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsTscSecuredSettings
    {
        #region <IDispatch>
        int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        HRESULT put_StartProgram(string pStartProgram);
        HRESULT get_StartProgram(out string pStartProgram);
        HRESULT put_WorkDir(string pWorkDir);
        HRESULT get_WorkDir(out string pWorkDir);
        HRESULT put_FullScreen(int pfFullScreen);
        HRESULT get_FullScreen(out int pfFullScreen);
    }

    [ComImport]
    [Guid("809945cc-4b3b-4a92-a6b0-dbf9b5f2ef2d")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsTscAdvancedSettings
    {
        #region <IDispatch>
        int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        HRESULT put_Compress(int pcompress);
        HRESULT get_Compress(out int pcompress);
        HRESULT put_BitmapPeristence(int pbitmapPeristence);
        HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        HRESULT put_KeyBoardLayoutStr(string _arg1);
        HRESULT put_PluginDlls(string _arg1);
        HRESULT put_IconFile(string _arg1);
        HRESULT put_IconIndex(int _arg1);
        HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        HRESULT put_DisableRdpdr(int pDisableRdpdr);
        HRESULT get_DisableRdpdr(out int pDisableRdpdr);
    }

    [ComImport]
    [Guid("92b4a539-7115-4b7c-a5a9-e5d9efc2780a")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient : IMsTscAx
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        HRESULT put_ColorDepth(int pcolorDepth);
        HRESULT get_ColorDepth(out int pcolorDepth);
        HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        HRESULT SetChannelOptions(string chanName, int chanOptions);
        HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        HRESULT RequestClose(out ControlCloseStatus pCloseStatus);
    }

    public enum ExtendedDisconnectReasonCode
    {
        exDiscReasonNoInfo = 0,
        exDiscReasonAPIInitiatedDisconnect = 1,
        exDiscReasonAPIInitiatedLogoff = 2,
        exDiscReasonServerIdleTimeout = 3,
        exDiscReasonServerLogonTimeout = 4,
        exDiscReasonReplacedByOtherConnection = 5,
        exDiscReasonOutOfMemory = 6,
        exDiscReasonServerDeniedConnection = 7,
        exDiscReasonServerDeniedConnectionFips = 8,
        exDiscReasonServerInsufficientPrivileges = 9,
        exDiscReasonServerFreshCredsRequired = 10,
        exDiscReasonRpcInitiatedDisconnectByUser = 11,
        exDiscReasonLogoffByUser = 12,
        exDiscReasonLicenseInternal = 256,
        exDiscReasonLicenseNoLicenseServer = 257,
        exDiscReasonLicenseNoLicense = 258,
        exDiscReasonLicenseErrClientMsg = 259,
        exDiscReasonLicenseHwidDoesntMatchLicense = 260,
        exDiscReasonLicenseErrClientLicense = 261,
        exDiscReasonLicenseCantFinishProtocol = 262,
        exDiscReasonLicenseClientEndedProtocol = 263,
        exDiscReasonLicenseErrClientEncryption = 264,
        exDiscReasonLicenseCantUpgradeLicense = 265,
        exDiscReasonLicenseNoRemoteConnections = 266,
        exDiscReasonLicenseCreatingLicStoreAccDenied = 267,
        exDiscReasonRdpEncInvalidCredentials = 768,
        exDiscReasonProtocolRangeStart = 4096,
        exDiscReasonProtocolRangeEnd = 32767
    };

    public enum ControlCloseStatus
    {
        controlCloseCanProceed = 0,
        controlCloseWaitForEvents = 1
    };

    [ComImport]
    [Guid("3c65b4ab-12b3-465b-acd4-b8dad3bff9e2")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings : IMsTscAdvancedSettings
    {
        #region <IDispatch>
       new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        HRESULT put_SmoothScroll(int psmoothScroll);
        HRESULT get_SmoothScroll(out int psmoothScroll);
        HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        HRESULT put_ShadowBitmap(int pshadowBitmap);
        HRESULT get_ShadowBitmap(out int pshadowBitmap);
        HRESULT put_TransportType(int ptransportType);
        HRESULT get_TransportType(out int ptransportType);
        HRESULT put_SasSequence(int psasSequence);
        HRESULT get_SasSequence(out int psasSequence);
        HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        HRESULT put_RDPPort(int prdpPort);
        HRESULT get_RDPPort(out int prdpPort);
        HRESULT put_EnableMouse(int penableMouse);
        HRESULT get_EnableMouse(out int penableMouse);
        HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        HRESULT put_EnableWindowsKey(int penableWindowsKey);
        HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        HRESULT put_MaximizeShell(int pmaximizeShell);
        HRESULT get_MaximizeShell(out int pmaximizeShell);
        HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        HRESULT put_HotKeyAltTab(int photKeyAltTab);
        HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        HRESULT put_PersistCacheDirectory(string _arg1);
        HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        HRESULT put_minInputSendInterval(int pminInputSendInterval);
        HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        HRESULT put_maxEventCount(int pmaxEventCount);
        HRESULT get_maxEventCount(out int pmaxEventCount);
        HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        HRESULT put_shutdownTimeout(int pshutdownTimeout);
        HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        HRESULT put_KeyboardType(int pkeyboardType);
        HRESULT get_KeyboardType(out int pkeyboardType);
        HRESULT put_KeyboardSubType(int pkeyboardSubType);
        HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        HRESULT put_BitmapPersistence(int pbitmapPersistence);
        HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        HRESULT put_ClearTextPassword(string _arg1);
        HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        HRESULT put_LoadBalanceInfo(string pLBInfo);
        HRESULT get_LoadBalanceInfo(out string pLBInfo);
        HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        HRESULT put_PerformanceFlags(int pDisableList);
        HRESULT get_PerformanceFlags(out int pDisableList);
        //HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);
    }

    [ComImport]
    [Guid("605befcf-39c1-45cc-a811-068fb7be346d")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientSecuredSettings : IMsTscSecuredSettings
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_StartProgram(string pStartProgram);
        new HRESULT get_StartProgram(out string pStartProgram);
        new HRESULT put_WorkDir(string pWorkDir);
        new HRESULT get_WorkDir(out string pWorkDir);
        new HRESULT put_FullScreen(int pfFullScreen);
        new HRESULT get_FullScreen(out int pfFullScreen);

        HRESULT put_KeyboardHookMode(int pkeyboardHookMode);
        HRESULT get_KeyboardHookMode(out int pkeyboardHookMode);
        HRESULT put_AudioRedirectionMode(int pAudioRedirectionMode);
        HRESULT get_AudioRedirectionMode(out int pAudioRedirectionMode);
    }

    [ComImport]
    [Guid("9ac42117-2b76-4320-aa44-0e616ab8437b")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings2 : IMsRdpClientAdvancedSettings
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        new HRESULT put_SmoothScroll(int psmoothScroll);
        new HRESULT get_SmoothScroll(out int psmoothScroll);
        new HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        new HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        new HRESULT put_ShadowBitmap(int pshadowBitmap);
        new HRESULT get_ShadowBitmap(out int pshadowBitmap);
        new HRESULT put_TransportType(int ptransportType);
        new HRESULT get_TransportType(out int ptransportType);
        new HRESULT put_SasSequence(int psasSequence);
        new HRESULT get_SasSequence(out int psasSequence);
        new HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        new HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        new HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        new HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        new HRESULT put_RDPPort(int prdpPort);
        new HRESULT get_RDPPort(out int prdpPort);
        new HRESULT put_EnableMouse(int penableMouse);
        new HRESULT get_EnableMouse(out int penableMouse);
        new HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        new HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        new HRESULT put_EnableWindowsKey(int penableWindowsKey);
        new HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        new HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        new HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        new HRESULT put_MaximizeShell(int pmaximizeShell);
        new HRESULT get_MaximizeShell(out int pmaximizeShell);
        new HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        new HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        new HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        new HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        new HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        new HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        new HRESULT put_HotKeyAltTab(int photKeyAltTab);
        new HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        new HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        new HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        new HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        new HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        new HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        new HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        new HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        new HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        new HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        new HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        new HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        new HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        new HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        new HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        new HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        new HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        new HRESULT put_PersistCacheDirectory(string _arg1);
        new HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        new HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        new HRESULT put_minInputSendInterval(int pminInputSendInterval);
        new HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        new HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        new HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        new HRESULT put_maxEventCount(int pmaxEventCount);
        new HRESULT get_maxEventCount(out int pmaxEventCount);
        new HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        new HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        new HRESULT put_shutdownTimeout(int pshutdownTimeout);
        new HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        new HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        new HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        new HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        new HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        new HRESULT put_KeyboardType(int pkeyboardType);
        new HRESULT get_KeyboardType(out int pkeyboardType);
        new HRESULT put_KeyboardSubType(int pkeyboardSubType);
        new HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        new HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        new HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        new HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        new HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        new HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        new HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        new HRESULT put_BitmapPersistence(int pbitmapPersistence);
        new HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        new HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        new HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        new HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        new HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        new HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        new HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        new HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        new HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        new HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        new HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        new HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        new HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT put_LoadBalanceInfo(string pLBInfo);
        new HRESULT get_LoadBalanceInfo(out string pLBInfo);
        new HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        new HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        new HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        new HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        new HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        new HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        new HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        new HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        new HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        new HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        new HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        new HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        new HRESULT put_PerformanceFlags(int pDisableList);
        new HRESULT get_PerformanceFlags(out int pDisableList);
        // new HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        new HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        new HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        new HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);

        HRESULT get_CanAutoReconnect(out VARIANT_BOOL pfCanAutoReconnect);
        HRESULT put_EnableAutoReconnect(VARIANT_BOOL pfEnableAutoReconnect);
        HRESULT get_EnableAutoReconnect(out VARIANT_BOOL pfEnableAutoReconnect);
        HRESULT put_MaxReconnectAttempts(int pMaxReconnectAttempts);
        HRESULT get_MaxReconnectAttempts(out int pMaxReconnectAttempts);
    }

    [ComImport]
    [Guid("19cd856b-c542-4c53-acee-f127e3be1a59")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings3 : IMsRdpClientAdvancedSettings2
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        new HRESULT put_SmoothScroll(int psmoothScroll);
        new HRESULT get_SmoothScroll(out int psmoothScroll);
        new HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        new HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        new HRESULT put_ShadowBitmap(int pshadowBitmap);
        new HRESULT get_ShadowBitmap(out int pshadowBitmap);
        new HRESULT put_TransportType(int ptransportType);
        new HRESULT get_TransportType(out int ptransportType);
        new HRESULT put_SasSequence(int psasSequence);
        new HRESULT get_SasSequence(out int psasSequence);
        new HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        new HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        new HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        new HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        new HRESULT put_RDPPort(int prdpPort);
        new HRESULT get_RDPPort(out int prdpPort);
        new HRESULT put_EnableMouse(int penableMouse);
        new HRESULT get_EnableMouse(out int penableMouse);
        new HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        new HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        new HRESULT put_EnableWindowsKey(int penableWindowsKey);
        new HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        new HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        new HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        new HRESULT put_MaximizeShell(int pmaximizeShell);
        new HRESULT get_MaximizeShell(out int pmaximizeShell);
        new HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        new HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        new HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        new HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        new HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        new HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        new HRESULT put_HotKeyAltTab(int photKeyAltTab);
        new HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        new HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        new HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        new HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        new HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        new HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        new HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        new HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        new HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        new HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        new HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        new HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        new HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        new HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        new HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        new HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        new HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        new HRESULT put_PersistCacheDirectory(string _arg1);
        new HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        new HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        new HRESULT put_minInputSendInterval(int pminInputSendInterval);
        new HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        new HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        new HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        new HRESULT put_maxEventCount(int pmaxEventCount);
        new HRESULT get_maxEventCount(out int pmaxEventCount);
        new HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        new HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        new HRESULT put_shutdownTimeout(int pshutdownTimeout);
        new HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        new HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        new HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        new HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        new HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        new HRESULT put_KeyboardType(int pkeyboardType);
        new HRESULT get_KeyboardType(out int pkeyboardType);
        new HRESULT put_KeyboardSubType(int pkeyboardSubType);
        new HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        new HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        new HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        new HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        new HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        new HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        new HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        new HRESULT put_BitmapPersistence(int pbitmapPersistence);
        new HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        new HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        new HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        new HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        new HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        new HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        new HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        new HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        new HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        new HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        new HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        new HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        new HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT put_LoadBalanceInfo(string pLBInfo);
        new HRESULT get_LoadBalanceInfo(out string pLBInfo);
        new HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        new HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        new HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        new HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        new HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        new HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        new HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        new HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        new HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        new HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        new HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        new HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        new HRESULT put_PerformanceFlags(int pDisableList);
        new HRESULT get_PerformanceFlags(out int pDisableList);
        // new HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        new HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        new HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        new HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);

        new HRESULT get_CanAutoReconnect(out VARIANT_BOOL pfCanAutoReconnect);
        new HRESULT put_EnableAutoReconnect(VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT get_EnableAutoReconnect(out VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT put_MaxReconnectAttempts(int pMaxReconnectAttempts);
        new HRESULT get_MaxReconnectAttempts(out int pMaxReconnectAttempts);

        HRESULT put_ConnectionBarShowMinimizeButton(VARIANT_BOOL pfShowMinimize);
        HRESULT get_ConnectionBarShowMinimizeButton(out VARIANT_BOOL pfShowMinimize);
        HRESULT put_ConnectionBarShowRestoreButton(VARIANT_BOOL pfShowRestore);
        HRESULT get_ConnectionBarShowRestoreButton(out VARIANT_BOOL pfShowRestore);
    }

    [ComImport]
    [Guid("fba7f64e-7345-4405-ae50-fa4a763dc0de")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings4 : IMsRdpClientAdvancedSettings3
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        new HRESULT put_SmoothScroll(int psmoothScroll);
        new HRESULT get_SmoothScroll(out int psmoothScroll);
        new HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        new HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        new HRESULT put_ShadowBitmap(int pshadowBitmap);
        new HRESULT get_ShadowBitmap(out int pshadowBitmap);
        new HRESULT put_TransportType(int ptransportType);
        new HRESULT get_TransportType(out int ptransportType);
        new HRESULT put_SasSequence(int psasSequence);
        new HRESULT get_SasSequence(out int psasSequence);
        new HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        new HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        new HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        new HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        new HRESULT put_RDPPort(int prdpPort);
        new HRESULT get_RDPPort(out int prdpPort);
        new HRESULT put_EnableMouse(int penableMouse);
        new HRESULT get_EnableMouse(out int penableMouse);
        new HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        new HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        new HRESULT put_EnableWindowsKey(int penableWindowsKey);
        new HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        new HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        new HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        new HRESULT put_MaximizeShell(int pmaximizeShell);
        new HRESULT get_MaximizeShell(out int pmaximizeShell);
        new HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        new HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        new HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        new HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        new HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        new HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        new HRESULT put_HotKeyAltTab(int photKeyAltTab);
        new HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        new HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        new HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        new HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        new HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        new HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        new HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        new HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        new HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        new HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        new HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        new HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        new HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        new HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        new HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        new HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        new HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        new HRESULT put_PersistCacheDirectory(string _arg1);
        new HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        new HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        new HRESULT put_minInputSendInterval(int pminInputSendInterval);
        new HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        new HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        new HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        new HRESULT put_maxEventCount(int pmaxEventCount);
        new HRESULT get_maxEventCount(out int pmaxEventCount);
        new HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        new HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        new HRESULT put_shutdownTimeout(int pshutdownTimeout);
        new HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        new HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        new HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        new HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        new HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        new HRESULT put_KeyboardType(int pkeyboardType);
        new HRESULT get_KeyboardType(out int pkeyboardType);
        new HRESULT put_KeyboardSubType(int pkeyboardSubType);
        new HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        new HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        new HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        new HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        new HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        new HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        new HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        new HRESULT put_BitmapPersistence(int pbitmapPersistence);
        new HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        new HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        new HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        new HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        new HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        new HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        new HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        new HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        new HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        new HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        new HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        new HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        new HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT put_LoadBalanceInfo(string pLBInfo);
        new HRESULT get_LoadBalanceInfo(out string pLBInfo);
        new HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        new HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        new HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        new HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        new HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        new HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        new HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        new HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        new HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        new HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        new HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        new HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        new HRESULT put_PerformanceFlags(int pDisableList);
        new HRESULT get_PerformanceFlags(out int pDisableList);
        // new HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        new HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        new HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        new HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);

        new HRESULT get_CanAutoReconnect(out VARIANT_BOOL pfCanAutoReconnect);
        new HRESULT put_EnableAutoReconnect(VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT get_EnableAutoReconnect(out VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT put_MaxReconnectAttempts(int pMaxReconnectAttempts);
        new HRESULT get_MaxReconnectAttempts(out int pMaxReconnectAttempts);

        new HRESULT put_ConnectionBarShowMinimizeButton(VARIANT_BOOL pfShowMinimize);
        new HRESULT get_ConnectionBarShowMinimizeButton(out VARIANT_BOOL pfShowMinimize);
        new HRESULT put_ConnectionBarShowRestoreButton(VARIANT_BOOL pfShowRestore);
        new HRESULT get_ConnectionBarShowRestoreButton(out VARIANT_BOOL pfShowRestore);

        HRESULT put_AuthenticationLevel(uint puiAuthLevel);
        HRESULT get_AuthenticationLevel(out uint puiAuthLevel);
    }

    [ComImport]
    [Guid("fba7f64e-6783-4405-da45-fa4a763dabd0")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings5 : IMsRdpClientAdvancedSettings4
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        new HRESULT put_SmoothScroll(int psmoothScroll);
        new HRESULT get_SmoothScroll(out int psmoothScroll);
        new HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        new HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        new HRESULT put_ShadowBitmap(int pshadowBitmap);
        new HRESULT get_ShadowBitmap(out int pshadowBitmap);
        new HRESULT put_TransportType(int ptransportType);
        new HRESULT get_TransportType(out int ptransportType);
        new HRESULT put_SasSequence(int psasSequence);
        new HRESULT get_SasSequence(out int psasSequence);
        new HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        new HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        new HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        new HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        new HRESULT put_RDPPort(int prdpPort);
        new HRESULT get_RDPPort(out int prdpPort);
        new HRESULT put_EnableMouse(int penableMouse);
        new HRESULT get_EnableMouse(out int penableMouse);
        new HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        new HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        new HRESULT put_EnableWindowsKey(int penableWindowsKey);
        new HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        new HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        new HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        new HRESULT put_MaximizeShell(int pmaximizeShell);
        new HRESULT get_MaximizeShell(out int pmaximizeShell);
        new HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        new HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        new HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        new HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        new HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        new HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        new HRESULT put_HotKeyAltTab(int photKeyAltTab);
        new HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        new HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        new HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        new HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        new HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        new HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        new HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        new HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        new HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        new HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        new HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        new HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        new HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        new HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        new HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        new HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        new HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        new HRESULT put_PersistCacheDirectory(string _arg1);
        new HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        new HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        new HRESULT put_minInputSendInterval(int pminInputSendInterval);
        new HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        new HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        new HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        new HRESULT put_maxEventCount(int pmaxEventCount);
        new HRESULT get_maxEventCount(out int pmaxEventCount);
        new HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        new HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        new HRESULT put_shutdownTimeout(int pshutdownTimeout);
        new HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        new HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        new HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        new HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        new HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        new HRESULT put_KeyboardType(int pkeyboardType);
        new HRESULT get_KeyboardType(out int pkeyboardType);
        new HRESULT put_KeyboardSubType(int pkeyboardSubType);
        new HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        new HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        new HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        new HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        new HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        new HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        new HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        new HRESULT put_BitmapPersistence(int pbitmapPersistence);
        new HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        new HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        new HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        new HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        new HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        new HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        new HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        new HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        new HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        new HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        new HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        new HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        new HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT put_LoadBalanceInfo(string pLBInfo);
        new HRESULT get_LoadBalanceInfo(out string pLBInfo);
        new HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        new HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        new HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        new HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        new HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        new HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        new HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        new HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        new HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        new HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        new HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        new HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        new HRESULT put_PerformanceFlags(int pDisableList);
        new HRESULT get_PerformanceFlags(out int pDisableList);
        // new HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        new HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        new HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        new HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);

        new HRESULT get_CanAutoReconnect(out VARIANT_BOOL pfCanAutoReconnect);
        new HRESULT put_EnableAutoReconnect(VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT get_EnableAutoReconnect(out VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT put_MaxReconnectAttempts(int pMaxReconnectAttempts);
        new HRESULT get_MaxReconnectAttempts(out int pMaxReconnectAttempts);

        new HRESULT put_ConnectionBarShowMinimizeButton(VARIANT_BOOL pfShowMinimize);
        new HRESULT get_ConnectionBarShowMinimizeButton(out VARIANT_BOOL pfShowMinimize);
        new HRESULT put_ConnectionBarShowRestoreButton(VARIANT_BOOL pfShowRestore);
        new HRESULT get_ConnectionBarShowRestoreButton(out VARIANT_BOOL pfShowRestore);

        new HRESULT put_AuthenticationLevel(uint puiAuthLevel);
        new HRESULT get_AuthenticationLevel(out uint puiAuthLevel);

        HRESULT put_RedirectClipboard(VARIANT_BOOL pfRedirectClipboard);
        HRESULT get_RedirectClipboard(out VARIANT_BOOL pfRedirectClipboard);
        HRESULT put_AudioRedirectionMode(uint puiAudioRedirectionMode);
        HRESULT get_AudioRedirectionMode(out uint puiAudioRedirectionMode);
        HRESULT put_ConnectionBarShowPinButton(VARIANT_BOOL pfShowPin);
        HRESULT get_ConnectionBarShowPinButton(out VARIANT_BOOL pfShowPin);
        HRESULT put_PublicMode(VARIANT_BOOL pfPublicMode);
        HRESULT get_PublicMode(out VARIANT_BOOL pfPublicMode);
        HRESULT put_RedirectDevices(VARIANT_BOOL pfRedirectPnPDevices);
        HRESULT get_RedirectDevices(out VARIANT_BOOL pfRedirectPnPDevices);
        HRESULT put_RedirectPOSDevices(VARIANT_BOOL pfRedirectPOSDevices);
        HRESULT get_RedirectPOSDevices(out VARIANT_BOOL pfRedirectPOSDevices);
        HRESULT put_BitmapVirtualCache32BppSize(int pBitmapVirtualCache32BppSize);
        HRESULT get_BitmapVirtualCache32BppSize(out int pBitmapVirtualCache32BppSize);
    }

    [ComImport]
    [Guid("222c4b5d-45d9-4df0-a7c6-60cf9089d285")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings6 : IMsRdpClientAdvancedSettings5
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        new HRESULT put_SmoothScroll(int psmoothScroll);
        new HRESULT get_SmoothScroll(out int psmoothScroll);
        new HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        new HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        new HRESULT put_ShadowBitmap(int pshadowBitmap);
        new HRESULT get_ShadowBitmap(out int pshadowBitmap);
        new HRESULT put_TransportType(int ptransportType);
        new HRESULT get_TransportType(out int ptransportType);
        new HRESULT put_SasSequence(int psasSequence);
        new HRESULT get_SasSequence(out int psasSequence);
        new HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        new HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        new HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        new HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        new HRESULT put_RDPPort(int prdpPort);
        new HRESULT get_RDPPort(out int prdpPort);
        new HRESULT put_EnableMouse(int penableMouse);
        new HRESULT get_EnableMouse(out int penableMouse);
        new HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        new HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        new HRESULT put_EnableWindowsKey(int penableWindowsKey);
        new HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        new HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        new HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        new HRESULT put_MaximizeShell(int pmaximizeShell);
        new HRESULT get_MaximizeShell(out int pmaximizeShell);
        new HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        new HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        new HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        new HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        new HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        new HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        new HRESULT put_HotKeyAltTab(int photKeyAltTab);
        new HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        new HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        new HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        new HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        new HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        new HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        new HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        new HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        new HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        new HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        new HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        new HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        new HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        new HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        new HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        new HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        new HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        new HRESULT put_PersistCacheDirectory(string _arg1);
        new HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        new HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        new HRESULT put_minInputSendInterval(int pminInputSendInterval);
        new HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        new HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        new HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        new HRESULT put_maxEventCount(int pmaxEventCount);
        new HRESULT get_maxEventCount(out int pmaxEventCount);
        new HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        new HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        new HRESULT put_shutdownTimeout(int pshutdownTimeout);
        new HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        new HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        new HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        new HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        new HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        new HRESULT put_KeyboardType(int pkeyboardType);
        new HRESULT get_KeyboardType(out int pkeyboardType);
        new HRESULT put_KeyboardSubType(int pkeyboardSubType);
        new HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        new HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        new HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        new HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        new HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        new HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        new HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        new HRESULT put_BitmapPersistence(int pbitmapPersistence);
        new HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        new HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        new HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        new HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        new HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        new HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        new HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        new HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        new HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        new HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        new HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        new HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        new HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT put_LoadBalanceInfo(string pLBInfo);
        new HRESULT get_LoadBalanceInfo(out string pLBInfo);
        new HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        new HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        new HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        new HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        new HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        new HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        new HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        new HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        new HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        new HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        new HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        new HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        new HRESULT put_PerformanceFlags(int pDisableList);
        new HRESULT get_PerformanceFlags(out int pDisableList);
        // new HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        new HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        new HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        new HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);

        new HRESULT get_CanAutoReconnect(out VARIANT_BOOL pfCanAutoReconnect);
        new HRESULT put_EnableAutoReconnect(VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT get_EnableAutoReconnect(out VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT put_MaxReconnectAttempts(int pMaxReconnectAttempts);
        new HRESULT get_MaxReconnectAttempts(out int pMaxReconnectAttempts);

        new HRESULT put_ConnectionBarShowMinimizeButton(VARIANT_BOOL pfShowMinimize);
        new HRESULT get_ConnectionBarShowMinimizeButton(out VARIANT_BOOL pfShowMinimize);
        new HRESULT put_ConnectionBarShowRestoreButton(VARIANT_BOOL pfShowRestore);
        new HRESULT get_ConnectionBarShowRestoreButton(out VARIANT_BOOL pfShowRestore);

        new HRESULT put_AuthenticationLevel(uint puiAuthLevel);
        new HRESULT get_AuthenticationLevel(out uint puiAuthLevel);

        new HRESULT put_RedirectClipboard(VARIANT_BOOL pfRedirectClipboard);
        new HRESULT get_RedirectClipboard(out VARIANT_BOOL pfRedirectClipboard);
        new HRESULT put_AudioRedirectionMode(uint puiAudioRedirectionMode);
        new HRESULT get_AudioRedirectionMode(out uint puiAudioRedirectionMode);
        new HRESULT put_ConnectionBarShowPinButton(VARIANT_BOOL pfShowPin);
        new HRESULT get_ConnectionBarShowPinButton(out VARIANT_BOOL pfShowPin);
        new HRESULT put_PublicMode(VARIANT_BOOL pfPublicMode);
        new HRESULT get_PublicMode(out VARIANT_BOOL pfPublicMode);
        new HRESULT put_RedirectDevices(VARIANT_BOOL pfRedirectPnPDevices);
        new HRESULT get_RedirectDevices(out VARIANT_BOOL pfRedirectPnPDevices);
        new HRESULT put_RedirectPOSDevices(VARIANT_BOOL pfRedirectPOSDevices);
        new HRESULT get_RedirectPOSDevices(out VARIANT_BOOL pfRedirectPOSDevices);
        new HRESULT put_BitmapVirtualCache32BppSize(int pBitmapVirtualCache32BppSize);
        new HRESULT get_BitmapVirtualCache32BppSize(out int pBitmapVirtualCache32BppSize);

        HRESULT put_RelativeMouseMode(VARIANT_BOOL pfRelativeMouseMode);
        HRESULT get_RelativeMouseMode(out VARIANT_BOOL pfRelativeMouseMode);
        HRESULT get_AuthenticationServiceClass(out string pbstrAuthServiceClass);
        HRESULT put_AuthenticationServiceClass(out string pbstrAuthServiceClass);
        HRESULT get_PCB(out string bstrPCB);
        HRESULT put_PCB(string bstrPCB);
        HRESULT put_HotKeyFocusReleaseLeft(int HotKeyFocusReleaseLeft);
        HRESULT get_HotKeyFocusReleaseLeft(out int HotKeyFocusReleaseLeft);
        HRESULT put_HotKeyFocusReleaseRight(int HotKeyFocusReleaseRight);
        HRESULT get_HotKeyFocusReleaseRight(out int HotKeyFocusReleaseRight);
        HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        HRESULT get_AuthenticationType(out uint puiAuthType);
        HRESULT put_ConnectToAdministerServer(VARIANT_BOOL pConnectToAdministerServer);
        HRESULT get_ConnectToAdministerServer(out VARIANT_BOOL pConnectToAdministerServer);
    }

    [ComImport]
    [Guid("e7e17dc4-3b71-4ba7-a8e6-281ffadca28f")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient2 : IMsRdpClient
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        HRESULT get_ConnectedStatusText(out string pConnectedStatusText);
    }

    [ComImport]
    [Guid("91b7cbc5-a72e-4fa0-9300-d647d7e897ff")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient3 : IMsRdpClient2
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);
    }

    [ComImport]
    [Guid("095e0738-d97d-488b-b9f6-dd0e8d66c0de")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient4 : IMsRdpClient3
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        new HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);

        HRESULT get_AdvancedSettings5(out IMsRdpClientAdvancedSettings4 ppAdvSettings);
    }

    [ComImport]
    [Guid("4eb5335b-6429-477d-b922-e06a28ecd8bf")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient5 : IMsRdpClient4
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        new HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);

        new HRESULT get_AdvancedSettings5(out IMsRdpClientAdvancedSettings4 ppAdvSettings);

        HRESULT get_TransportSettings(out IMsRdpClientTransportSettings ppXportSet);
        HRESULT get_AdvancedSettings6(out IMsRdpClientAdvancedSettings5 ppAdvSettings);
        HRESULT GetErrorDescription(uint disconnectReason, uint ExtendedDisconnectReason, out string pBstrErrorMsg);
        HRESULT get_RemoteProgram(out ITSRemoteProgram ppRemoteProgram);
        HRESULT get_MsRdpClientShell(out IMsRdpClientShell ppLauncher);
    }

    [ComImport]
    [Guid("720298c0-a099-46f5-9f82-96921bae4701")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientTransportSettings
    {
        #region <IDispatch>
        int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        HRESULT put_GatewayHostname(string pProxyHostname);
        HRESULT get_GatewayHostname(out string pProxyHostname);
        HRESULT put_GatewayUsageMethod(uint pulProxyUsageMethod);
        HRESULT get_GatewayUsageMethod(out uint pulProxyUsageMethod);
        HRESULT put_GatewayProfileUsageMethod(uint pulProxyProfileUsageMethod);
        HRESULT get_GatewayProfileUsageMethod(out uint pulProxyProfileUsageMethod);
        HRESULT put_GatewayCredsSource(uint pulProxyCredsSource);
        HRESULT get_GatewayCredsSource(out uint pulProxyCredsSource);
        HRESULT put_GatewayUserSelectedCredsSource(uint pulProxyCredsSource);
        HRESULT get_GatewayUserSelectedCredsSource(out uint pulProxyCredsSource);
        HRESULT get_GatewayIsSupported(out int pfProxyIsSupported);
        HRESULT get_GatewayDefaultUsageMethod(out uint pulProxyDefaultUsageMethod);
    }

    [ComImport]
    [Guid("fdd029f9-467a-4c49-8529-64b521dbd1b4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITSRemoteProgram
    {
        #region <IDispatch>
        int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        HRESULT put_RemoteProgramMode(VARIANT_BOOL pvboolRemoteProgramMode);
        HRESULT get_RemoteProgramMode(out VARIANT_BOOL pvboolRemoteProgramMode);
        [PreserveSig]
        HRESULT ServerStartProgram(string bstrExecutablePath, string bstrFilePath, string bstrWorkingDirectory, 
            VARIANT_BOOL vbExpandEnvVarInWorkingDirectoryOnServer, string bstrArguments, VARIANT_BOOL vbExpandEnvVarInArgumentsOnServer);
    }

    [ComImport]
    [Guid("d012ae6d-c19a-4bfe-b367-201f8911f134")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientShell
    {
        #region <IDispatch>
        int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        HRESULT Launch();
        HRESULT put_RdpFileContents(string pszRdpFile);
        HRESULT get_RdpFileContents(out string pszRdpFile);
        //HRESULT SetRdpProperty(string szProperty, VARIANT Value);
        HRESULT SetRdpProperty(string szProperty, IntPtr Value);
        //HRESULT GetRdpProperty(string szProperty, out VARIANT pValue);
        HRESULT GetRdpProperty(string szProperty, out IntPtr pValue);
        HRESULT get_IsRemoteProgramClientInstalled(out VARIANT_BOOL pbClientInstalled);
        HRESULT put_PublicMode(VARIANT_BOOL pfPublicMode);
        HRESULT get_PublicMode(out VARIANT_BOOL pfPublicMode);
        HRESULT ShowTrustedSitesManagementDialog();
    }

    [ComImport]
    [Guid("d43b7d80-8517-4b6d-9eac-96ad6800d7f2")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient6 : IMsRdpClient5
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        new HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);

        new HRESULT get_AdvancedSettings5(out IMsRdpClientAdvancedSettings4 ppAdvSettings);

        new HRESULT get_TransportSettings(out IMsRdpClientTransportSettings ppXportSet);
        new HRESULT get_AdvancedSettings6(out IMsRdpClientAdvancedSettings5 ppAdvSettings);
        new HRESULT GetErrorDescription(uint disconnectReason, uint ExtendedDisconnectReason, out string pBstrErrorMsg);
        new HRESULT get_RemoteProgram(out ITSRemoteProgram ppRemoteProgram);
        new HRESULT get_MsRdpClientShell(out IMsRdpClientShell ppLauncher);

        HRESULT get_AdvancedSettings7(out IMsRdpClientAdvancedSettings6 ppAdvSettings);
        HRESULT get_TransportSettings2(out IMsRdpClientTransportSettings2 ppXportSet2);
    }

    [ComImport]
    [Guid("67341688-d606-4c73-a5d2-2e0489009319")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientTransportSettings2 : IMsRdpClientTransportSettings
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_GatewayHostname(string pProxyHostname);
        new HRESULT get_GatewayHostname(out string pProxyHostname);
        new HRESULT put_GatewayUsageMethod(uint pulProxyUsageMethod);
        new HRESULT get_GatewayUsageMethod(out uint pulProxyUsageMethod);
        new HRESULT put_GatewayProfileUsageMethod(uint pulProxyProfileUsageMethod);
        new HRESULT get_GatewayProfileUsageMethod(out uint pulProxyProfileUsageMethod);
        new HRESULT put_GatewayCredsSource(uint pulProxyCredsSource);
        new HRESULT get_GatewayCredsSource(out uint pulProxyCredsSource);
        new HRESULT put_GatewayUserSelectedCredsSource(uint pulProxyCredsSource);
        new HRESULT get_GatewayUserSelectedCredsSource(out uint pulProxyCredsSource);
        new HRESULT get_GatewayIsSupported(out int pfProxyIsSupported);
        new HRESULT get_GatewayDefaultUsageMethod(out uint pulProxyDefaultUsageMethod);

        HRESULT put_GatewayCredSharing(uint pulProxyCredSharing);
        HRESULT get_GatewayCredSharing(out uint pulProxyCredSharing);
        HRESULT put_GatewayPreAuthRequirement(uint pulProxyPreAuthRequirement);
        HRESULT get_GatewayPreAuthRequirement(out uint pulProxyPreAuthRequirement);
        HRESULT put_GatewayPreAuthServerAddr(string pstringProxyPreAuthServerAddr);
        HRESULT get_GatewayPreAuthServerAddr(out string pstringProxyPreAuthServerAddr);
        HRESULT put_GatewaySupportUrl(string pstringProxySupportUrl);
        HRESULT get_GatewaySupportUrl(out string pstringProxySupportUrl);
        HRESULT put_GatewayEncryptedOtpCookie(string pstringEncryptedOtpCookie);
        HRESULT get_GatewayEncryptedOtpCookie(out string pstringEncryptedOtpCookie);
        HRESULT put_GatewayEncryptedOtpCookieSize(uint pulEncryptedOtpCookieSize);
        HRESULT get_GatewayEncryptedOtpCookieSize(out uint pulEncryptedOtpCookieSize);
        HRESULT put_GatewayUsername(string pProxyUsername);
        HRESULT get_GatewayUsername(out string pProxyUsername);
        HRESULT put_GatewayDomain(string pProxyDomain);
        HRESULT get_GatewayDomain(out string pProxyDomain);
        HRESULT put_GatewayPassword(string _arg1);
    }

    [ComImport]
    [Guid("b2a5b5ce-3461-444a-91d4-add26d070638")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient7 : IMsRdpClient6
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        [PreserveSig]
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        new HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);

        new HRESULT get_AdvancedSettings5(out IMsRdpClientAdvancedSettings4 ppAdvSettings);

        new HRESULT get_TransportSettings(out IMsRdpClientTransportSettings ppXportSet);
        new HRESULT get_AdvancedSettings6(out IMsRdpClientAdvancedSettings5 ppAdvSettings);
        new HRESULT GetErrorDescription(uint disconnectReason, uint ExtendedDisconnectReason, out string pBstrErrorMsg);
        new HRESULT get_RemoteProgram(out ITSRemoteProgram ppRemoteProgram);
        new HRESULT get_MsRdpClientShell(out IMsRdpClientShell ppLauncher);

        new HRESULT get_AdvancedSettings7(out IMsRdpClientAdvancedSettings6 ppAdvSettings);
        new HRESULT get_TransportSettings2(out IMsRdpClientTransportSettings2 ppXportSet2);

        HRESULT get_AdvancedSettings8(out IMsRdpClientAdvancedSettings7 ppAdvSettings);
        HRESULT get_TransportSettings3(out IMsRdpClientTransportSettings3 ppXportSet3);
        HRESULT GetStatusText(uint statusCode, out string pBstrStatusText);
        HRESULT get_SecuredSettings3(out IMsRdpClientSecuredSettings2 ppSecuredSettings);
        HRESULT get_RemoteProgram2(out ITSRemoteProgram2 ppRemoteProgram);
    }

    [ComImport]
    [Guid("4247e044-9271-43a9-bc49-e2ad9e855d62")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient8 : IMsRdpClient7
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        [PreserveSig]
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        new HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);

        new HRESULT get_AdvancedSettings5(out IMsRdpClientAdvancedSettings4 ppAdvSettings);

        new HRESULT get_TransportSettings(out IMsRdpClientTransportSettings ppXportSet);
        new HRESULT get_AdvancedSettings6(out IMsRdpClientAdvancedSettings5 ppAdvSettings);
        new HRESULT GetErrorDescription(uint disconnectReason, uint ExtendedDisconnectReason, out string pBstrErrorMsg);
        new HRESULT get_RemoteProgram(out ITSRemoteProgram ppRemoteProgram);
        new HRESULT get_MsRdpClientShell(out IMsRdpClientShell ppLauncher);

        new HRESULT get_AdvancedSettings7(out IMsRdpClientAdvancedSettings6 ppAdvSettings);
        new HRESULT get_TransportSettings2(out IMsRdpClientTransportSettings2 ppXportSet2);

        new HRESULT get_AdvancedSettings8(out IMsRdpClientAdvancedSettings7 ppAdvSettings);
        new HRESULT get_TransportSettings3(out IMsRdpClientTransportSettings3 ppXportSet3);
        new HRESULT GetStatusText(uint statusCode, out string pBstrStatusText);
        new HRESULT get_SecuredSettings3(out IMsRdpClientSecuredSettings2 ppSecuredSettings);
        new HRESULT get_RemoteProgram2(out ITSRemoteProgram2 ppRemoteProgram);

        HRESULT SendRemoteAction(RemoteSessionActionType actionType);
        HRESULT get_AdvancedSettings9(out IMsRdpClientAdvancedSettings8 ppAdvSettings);
        HRESULT Reconnect(uint ulWidth, uint ulHeight, out ControlReconnectStatus pReconnectStatus);
    }

    public enum RemoteSessionActionType
    {
        RemoteSessionActionCharms = 0,
        RemoteSessionActionAppbar = 1,
        RemoteSessionActionSnap = 2,
        RemoteSessionActionStartScreen = 3,
        RemoteSessionActionAppSwitch = 4,
        RemoteSessionActionActionCenter = 5
    };

    public enum ControlReconnectStatus
    {
        controlReconnectStarted = 0,
        controlReconnectBlocked = 1
    };

    [ComImport]
    [Guid("28904001-04b6-436c-a55b-0af1a0883dc9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient9 : IMsRdpClient8
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        [PreserveSig]
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        new HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);

        new HRESULT get_AdvancedSettings5(out IMsRdpClientAdvancedSettings4 ppAdvSettings);

        new HRESULT get_TransportSettings(out IMsRdpClientTransportSettings ppXportSet);
        new HRESULT get_AdvancedSettings6(out IMsRdpClientAdvancedSettings5 ppAdvSettings);
        new HRESULT GetErrorDescription(uint disconnectReason, uint ExtendedDisconnectReason, out string pBstrErrorMsg);
        new HRESULT get_RemoteProgram(out ITSRemoteProgram ppRemoteProgram);
        new HRESULT get_MsRdpClientShell(out IMsRdpClientShell ppLauncher);

        new HRESULT get_AdvancedSettings7(out IMsRdpClientAdvancedSettings6 ppAdvSettings);
        new HRESULT get_TransportSettings2(out IMsRdpClientTransportSettings2 ppXportSet2);

        new HRESULT get_AdvancedSettings8(out IMsRdpClientAdvancedSettings7 ppAdvSettings);
        new HRESULT get_TransportSettings3(out IMsRdpClientTransportSettings3 ppXportSet3);
        new HRESULT GetStatusText(uint statusCode, out string pBstrStatusText);
        new HRESULT get_SecuredSettings3(out IMsRdpClientSecuredSettings2 ppSecuredSettings);
        new HRESULT get_RemoteProgram2(out ITSRemoteProgram2 ppRemoteProgram);

        new HRESULT SendRemoteAction(RemoteSessionActionType actionType);
        new HRESULT get_AdvancedSettings9(out IMsRdpClientAdvancedSettings8 ppAdvSettings);
        new HRESULT Reconnect(uint ulWidth, uint ulHeight, out ControlReconnectStatus pReconnectStatus);

        HRESULT get_TransportSettings4(out IMsRdpClientTransportSettings4 ppXportSet4);
        HRESULT SyncSessionDisplaySettings();
        HRESULT UpdateSessionDisplaySettings(uint ulDesktopWidth, uint ulDesktopHeight, uint ulPhysicalWidth, uint ulPhysicalHeight, uint ulOrientation, uint ulDesktopScaleFactor, uint ulDeviceScaleFactor);
        HRESULT attachEvent(string eventName, IntPtr callback);
        HRESULT detachEvent(string eventName, IntPtr callback);
    }

    [ComImport]
    [Guid("7ed92c39-eb38-4927-a70a-708ac5a59321")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClient10 : IMsRdpClient9
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        [PreserveSig]
        new HRESULT put_Server(string pServer);
        new HRESULT get_Server(out string pServer);
        [PreserveSig]
        new HRESULT put_Domain(string pDomain);
        new HRESULT get_Domain(out string pDomain);
        [PreserveSig]
        new HRESULT put_UserName(string pUserName);
        new HRESULT get_UserName(out string pUserName);
        new HRESULT put_DisconnectedText(string pDisconnectedText);
        new HRESULT get_DisconnectedText(out string pDisconnectedText);
        new HRESULT put_ConnectingText(string pConnectingText);
        new HRESULT get_ConnectingText(out string pConnectingText);
        new HRESULT get_Connected(out short pIsConnected);
        [PreserveSig]
        new HRESULT put_DesktopWidth(int pVal);
        new HRESULT get_DesktopWidth(out int pVal);
        [PreserveSig]
        new HRESULT put_DesktopHeight(int pVal);
        new HRESULT get_DesktopHeight(out int pVal);
        new HRESULT put_StartConnected(int pfStartConnected);
        new HRESULT get_StartConnected(out int pfStartConnected);
        new HRESULT get_HorizontalScrollBarVisible(out int pfHScrollVisible);
        new HRESULT get_VerticalScrollBarVisible(out int pfVScrollVisible);
        new HRESULT put_FullScreenTitle(string _arg1);
        new HRESULT get_CipherStrength(out int pCipherStrength);
        new HRESULT get_Version(out string pVersion);
        new HRESULT get_SecuredSettingsEnabled(out int pSecuredSettingsEnabled);
        new HRESULT get_SecuredSettings(out IMsTscSecuredSettings ppSecuredSettings);
        new HRESULT get_AdvancedSettings(out IMsTscAdvancedSettings ppAdvSettings);
        //new HRESULT get_Debugger(out IMsTscDebug ppDebugger);
        new HRESULT get_Debugger(out IntPtr ppDebugger);
        [PreserveSig]
        new HRESULT Connect();
        new HRESULT Disconnect();
        new HRESULT CreateChannels(string newVal);
        new HRESULT SendOnChannel(string chanName, string ChanData);

        new HRESULT put_ColorDepth(int pcolorDepth);
        new HRESULT get_ColorDepth(out int pcolorDepth);
        new HRESULT get_AdvancedSettings2(out IMsRdpClientAdvancedSettings ppAdvSettings);
        new HRESULT get_SecuredSettings2(out IMsRdpClientSecuredSettings ppSecuredSettings);
        new HRESULT get_ExtendedDisconnectReason(out ExtendedDisconnectReasonCode pExtendedDisconnectReason);
        new HRESULT put_FullScreen(VARIANT_BOOL pfFullScreen);
        [PreserveSig]
        new HRESULT get_FullScreen(out VARIANT_BOOL pfFullScreen);
        new HRESULT SetChannelOptions(string chanName, int chanOptions);
        new HRESULT GetChannelOptions(string chanName, out int pChanOptions);
        new HRESULT RequestClose(out ControlCloseStatus pCloseStatus);

        new HRESULT get_AdvancedSettings3(out IMsRdpClientAdvancedSettings2 ppAdvSettings);
        new HRESULT put_ConnectedStatusText(string pConnectedStatusText);
        new HRESULT get_ConnectedStatusText(out string pConnectedStatusText);

        new HRESULT get_AdvancedSettings4(out IMsRdpClientAdvancedSettings3 ppAdvSettings);

        new HRESULT get_AdvancedSettings5(out IMsRdpClientAdvancedSettings4 ppAdvSettings);

        new HRESULT get_TransportSettings(out IMsRdpClientTransportSettings ppXportSet);
        new HRESULT get_AdvancedSettings6(out IMsRdpClientAdvancedSettings5 ppAdvSettings);
        new HRESULT GetErrorDescription(uint disconnectReason, uint ExtendedDisconnectReason, out string pBstrErrorMsg);
        new HRESULT get_RemoteProgram(out ITSRemoteProgram ppRemoteProgram);
        new HRESULT get_MsRdpClientShell(out IMsRdpClientShell ppLauncher);

        new HRESULT get_AdvancedSettings7(out IMsRdpClientAdvancedSettings6 ppAdvSettings);
        new HRESULT get_TransportSettings2(out IMsRdpClientTransportSettings2 ppXportSet2);

        new HRESULT get_AdvancedSettings8(out IMsRdpClientAdvancedSettings7 ppAdvSettings);
        new HRESULT get_TransportSettings3(out IMsRdpClientTransportSettings3 ppXportSet3);
        new HRESULT GetStatusText(uint statusCode, out string pBstrStatusText);
        new HRESULT get_SecuredSettings3(out IMsRdpClientSecuredSettings2 ppSecuredSettings);
        new HRESULT get_RemoteProgram2(out ITSRemoteProgram2 ppRemoteProgram);

        new HRESULT SendRemoteAction(RemoteSessionActionType actionType);
        new HRESULT get_AdvancedSettings9(out IMsRdpClientAdvancedSettings8 ppAdvSettings);
        new HRESULT Reconnect(uint ulWidth, uint ulHeight, out ControlReconnectStatus pReconnectStatus);

        new HRESULT get_TransportSettings4(out IMsRdpClientTransportSettings4 ppXportSet4);
        new HRESULT SyncSessionDisplaySettings();
        new HRESULT UpdateSessionDisplaySettings(uint ulDesktopWidth, uint ulDesktopHeight, uint ulPhysicalWidth, uint ulPhysicalHeight, uint ulOrientation, uint ulDesktopScaleFactor, uint ulDeviceScaleFactor);
        new HRESULT attachEvent(string eventName, IntPtr callback);
        new HRESULT detachEvent(string eventName, IntPtr callback);

        HRESULT get_RemoteProgram3(out ITSRemoteProgram3 ppRemoteProgram);
    }







    [ComImport]
    [Guid("26036036-4010-4578-8091-0db9a1edf9c3")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings7 : IMsRdpClientAdvancedSettings6
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        new HRESULT put_SmoothScroll(int psmoothScroll);
        new HRESULT get_SmoothScroll(out int psmoothScroll);
        new HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        new HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        new HRESULT put_ShadowBitmap(int pshadowBitmap);
        new HRESULT get_ShadowBitmap(out int pshadowBitmap);
        new HRESULT put_TransportType(int ptransportType);
        new HRESULT get_TransportType(out int ptransportType);
        new HRESULT put_SasSequence(int psasSequence);
        new HRESULT get_SasSequence(out int psasSequence);
        new HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        new HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        new HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        new HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        new HRESULT put_RDPPort(int prdpPort);
        new HRESULT get_RDPPort(out int prdpPort);
        new HRESULT put_EnableMouse(int penableMouse);
        new HRESULT get_EnableMouse(out int penableMouse);
        new HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        new HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        new HRESULT put_EnableWindowsKey(int penableWindowsKey);
        new HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        new HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        new HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        new HRESULT put_MaximizeShell(int pmaximizeShell);
        new HRESULT get_MaximizeShell(out int pmaximizeShell);
        new HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        new HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        new HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        new HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        new HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        new HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        new HRESULT put_HotKeyAltTab(int photKeyAltTab);
        new HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        new HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        new HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        new HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        new HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        new HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        new HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        new HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        new HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        new HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        new HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        new HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        new HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        new HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        new HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        new HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        new HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        new HRESULT put_PersistCacheDirectory(string _arg1);
        new HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        new HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        new HRESULT put_minInputSendInterval(int pminInputSendInterval);
        new HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        new HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        new HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        new HRESULT put_maxEventCount(int pmaxEventCount);
        new HRESULT get_maxEventCount(out int pmaxEventCount);
        new HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        new HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        new HRESULT put_shutdownTimeout(int pshutdownTimeout);
        new HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        new HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        new HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        new HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        new HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        new HRESULT put_KeyboardType(int pkeyboardType);
        new HRESULT get_KeyboardType(out int pkeyboardType);
        new HRESULT put_KeyboardSubType(int pkeyboardSubType);
        new HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        new HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        new HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        new HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        new HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        new HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        new HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        new HRESULT put_BitmapPersistence(int pbitmapPersistence);
        new HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        new HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        new HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        new HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        new HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        new HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        new HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        new HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        new HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        new HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        new HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        new HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        new HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT put_LoadBalanceInfo(string pLBInfo);
        new HRESULT get_LoadBalanceInfo(out string pLBInfo);
        new HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        new HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        new HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        new HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        new HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        new HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        new HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        new HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        new HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        new HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        new HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        new HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        new HRESULT put_PerformanceFlags(int pDisableList);
        new HRESULT get_PerformanceFlags(out int pDisableList);
        // new HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        new HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        new HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        new HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);

        new HRESULT get_CanAutoReconnect(out VARIANT_BOOL pfCanAutoReconnect);
        new HRESULT put_EnableAutoReconnect(VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT get_EnableAutoReconnect(out VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT put_MaxReconnectAttempts(int pMaxReconnectAttempts);
        new HRESULT get_MaxReconnectAttempts(out int pMaxReconnectAttempts);

        new HRESULT put_ConnectionBarShowMinimizeButton(VARIANT_BOOL pfShowMinimize);
        new HRESULT get_ConnectionBarShowMinimizeButton(out VARIANT_BOOL pfShowMinimize);
        new HRESULT put_ConnectionBarShowRestoreButton(VARIANT_BOOL pfShowRestore);
        new HRESULT get_ConnectionBarShowRestoreButton(out VARIANT_BOOL pfShowRestore);

        new HRESULT put_AuthenticationLevel(uint puiAuthLevel);
        new HRESULT get_AuthenticationLevel(out uint puiAuthLevel);

        new HRESULT put_RedirectClipboard(VARIANT_BOOL pfRedirectClipboard);
        new HRESULT get_RedirectClipboard(out VARIANT_BOOL pfRedirectClipboard);
        new HRESULT put_AudioRedirectionMode(uint puiAudioRedirectionMode);
        new HRESULT get_AudioRedirectionMode(out uint puiAudioRedirectionMode);
        new HRESULT put_ConnectionBarShowPinButton(VARIANT_BOOL pfShowPin);
        new HRESULT get_ConnectionBarShowPinButton(out VARIANT_BOOL pfShowPin);
        new HRESULT put_PublicMode(VARIANT_BOOL pfPublicMode);
        new HRESULT get_PublicMode(out VARIANT_BOOL pfPublicMode);
        new HRESULT put_RedirectDevices(VARIANT_BOOL pfRedirectPnPDevices);
        new HRESULT get_RedirectDevices(out VARIANT_BOOL pfRedirectPnPDevices);
        new HRESULT put_RedirectPOSDevices(VARIANT_BOOL pfRedirectPOSDevices);
        new HRESULT get_RedirectPOSDevices(out VARIANT_BOOL pfRedirectPOSDevices);
        new HRESULT put_BitmapVirtualCache32BppSize(int pBitmapVirtualCache32BppSize);
        new HRESULT get_BitmapVirtualCache32BppSize(out int pBitmapVirtualCache32BppSize);

        new HRESULT put_RelativeMouseMode(VARIANT_BOOL pfRelativeMouseMode);
        new HRESULT get_RelativeMouseMode(out VARIANT_BOOL pfRelativeMouseMode);
        new HRESULT get_AuthenticationServiceClass(out string pbstrAuthServiceClass);
        new HRESULT put_AuthenticationServiceClass(out string pbstrAuthServiceClass);
        new HRESULT get_PCB(out string bstrPCB);
        new HRESULT put_PCB(string bstrPCB);
        new HRESULT put_HotKeyFocusReleaseLeft(int HotKeyFocusReleaseLeft);
        new HRESULT get_HotKeyFocusReleaseLeft(out int HotKeyFocusReleaseLeft);
        new HRESULT put_HotKeyFocusReleaseRight(int HotKeyFocusReleaseRight);
        new HRESULT get_HotKeyFocusReleaseRight(out int HotKeyFocusReleaseRight);
        new HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        new HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        new HRESULT get_AuthenticationType(out uint puiAuthType);
        new HRESULT put_ConnectToAdministerServer(VARIANT_BOOL pConnectToAdministerServer);
        new HRESULT get_ConnectToAdministerServer(out VARIANT_BOOL pConnectToAdministerServer);

        HRESULT put_AudioCaptureRedirectionMode(VARIANT_BOOL pfRedir);
        HRESULT get_AudioCaptureRedirectionMode(out VARIANT_BOOL pfRedir);
        HRESULT put_VideoPlaybackMode(uint pVideoPlaybackMode);
        HRESULT get_VideoPlaybackMode(out uint pVideoPlaybackMode);
        HRESULT put_EnableSuperPan(VARIANT_BOOL pfEnableSuperPan);
        HRESULT get_EnableSuperPan(out VARIANT_BOOL pfEnableSuperPan);
        HRESULT put_SuperPanAccelerationFactor(uint puAccelFactor);
        HRESULT get_SuperPanAccelerationFactor(out uint puAccelFactor);
        HRESULT put_NegotiateSecurityLayer(VARIANT_BOOL pfNegotiate);
        HRESULT get_NegotiateSecurityLayer(out VARIANT_BOOL pfNegotiate);
        HRESULT put_AudioQualityMode(uint pAudioQualityMode);
        HRESULT get_AudioQualityMode(out uint pAudioQualityMode);
        HRESULT put_RedirectDirectX(VARIANT_BOOL pfRedirectDirectX);
        HRESULT get_RedirectDirectX(out VARIANT_BOOL pfRedirectDirectX);
        HRESULT put_NetworkConnectionType(uint pConnectionType);
        HRESULT get_NetworkConnectionType(out uint pConnectionType);
    }

    [ComImport]
    [Guid("89acb528-2557-4d16-8625-226a30e97e9a")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientAdvancedSettings8 : IMsRdpClientAdvancedSettings7
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_Compress(int pcompress);
        new HRESULT get_Compress(out int pcompress);
        new HRESULT put_BitmapPeristence(int pbitmapPeristence);
        new HRESULT get_BitmapPeristence(out int pbitmapPeristence);
        new HRESULT put_allowBackgroundInput(int pallowBackgroundInput);
        new HRESULT get_allowBackgroundInput(out int pallowBackgroundInput);
        new HRESULT put_KeyBoardLayoutStr(string _arg1);
        new HRESULT put_PluginDlls(string _arg1);
        new HRESULT put_IconFile(string _arg1);
        new HRESULT put_IconIndex(int _arg1);
        new HRESULT put_ContainerHandledFullScreen(int pContainerHandledFullScreen);
        new HRESULT get_ContainerHandledFullScreen(out int pContainerHandledFullScreen);
        new HRESULT put_DisableRdpdr(int pDisableRdpdr);
        new HRESULT get_DisableRdpdr(out int pDisableRdpdr);

        new HRESULT put_SmoothScroll(int psmoothScroll);
        new HRESULT get_SmoothScroll(out int psmoothScroll);
        new HRESULT put_AcceleratorPassthrough(int pacceleratorPassthrough);
        new HRESULT get_AcceleratorPassthrough(out int pacceleratorPassthrough);
        new HRESULT put_ShadowBitmap(int pshadowBitmap);
        new HRESULT get_ShadowBitmap(out int pshadowBitmap);
        new HRESULT put_TransportType(int ptransportType);
        new HRESULT get_TransportType(out int ptransportType);
        new HRESULT put_SasSequence(int psasSequence);
        new HRESULT get_SasSequence(out int psasSequence);
        new HRESULT put_EncryptionEnabled(int pencryptionEnabled);
        new HRESULT get_EncryptionEnabled(out int pencryptionEnabled);
        new HRESULT put_DedicatedTerminal(int pdedicatedTerminal);
        new HRESULT get_DedicatedTerminal(out int pdedicatedTerminal);
        new HRESULT put_RDPPort(int prdpPort);
        new HRESULT get_RDPPort(out int prdpPort);
        new HRESULT put_EnableMouse(int penableMouse);
        new HRESULT get_EnableMouse(out int penableMouse);
        new HRESULT put_DisableCtrlAltDel(int pdisableCtrlAltDel);
        new HRESULT get_DisableCtrlAltDel(out int pdisableCtrlAltDel);
        new HRESULT put_EnableWindowsKey(int penableWindowsKey);
        new HRESULT get_EnableWindowsKey(out int penableWindowsKey);
        new HRESULT put_DoubleClickDetect(int pdoubleClickDetect);
        new HRESULT get_DoubleClickDetect(out int pdoubleClickDetect);
        new HRESULT put_MaximizeShell(int pmaximizeShell);
        new HRESULT get_MaximizeShell(out int pmaximizeShell);
        new HRESULT put_HotKeyFullScreen(int photKeyFullScreen);
        new HRESULT get_HotKeyFullScreen(out int photKeyFullScreen);
        new HRESULT put_HotKeyCtrlEsc(int photKeyCtrlEsc);
        new HRESULT get_HotKeyCtrlEsc(out int photKeyCtrlEsc);
        new HRESULT put_HotKeyAltEsc(int photKeyAltEsc);
        new HRESULT get_HotKeyAltEsc(out int photKeyAltEsc);
        new HRESULT put_HotKeyAltTab(int photKeyAltTab);
        new HRESULT get_HotKeyAltTab(out int photKeyAltTab);
        new HRESULT put_HotKeyAltShiftTab(int photKeyAltShiftTab);
        new HRESULT get_HotKeyAltShiftTab(out int photKeyAltShiftTab);
        new HRESULT put_HotKeyAltSpace(int photKeyAltSpace);
        new HRESULT get_HotKeyAltSpace(out int photKeyAltSpace);
        new HRESULT put_HotKeyCtrlAltDel(int photKeyCtrlAltDel);
        new HRESULT get_HotKeyCtrlAltDel(out int photKeyCtrlAltDel);
        new HRESULT put_orderDrawThreshold(int porderDrawThreshold);
        new HRESULT get_orderDrawThreshold(out int porderDrawThreshold);
        new HRESULT put_BitmapCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapCacheSize(out int pbitmapCacheSize);
        new HRESULT put_BitmapVirtualCacheSize(int pbitmapCacheSize);
        new HRESULT get_BitmapVirtualCacheSize(out int pbitmapCacheSize);
        new HRESULT put_ScaleBitmapCachesByBPP(int pbScale);
        new HRESULT get_ScaleBitmapCachesByBPP(out int pbScale);
        new HRESULT put_NumBitmapCaches(int pnumBitmapCaches);
        new HRESULT get_NumBitmapCaches(out int pnumBitmapCaches);
        new HRESULT put_CachePersistenceActive(int pcachePersistenceActive);
        new HRESULT get_CachePersistenceActive(out int pcachePersistenceActive);
        new HRESULT put_PersistCacheDirectory(string _arg1);
        new HRESULT put_brushSupportLevel(int pbrushSupportLevel);
        new HRESULT get_brushSupportLevel(out int pbrushSupportLevel);
        new HRESULT put_minInputSendInterval(int pminInputSendInterval);
        new HRESULT get_minInputSendInterval(out int pminInputSendInterval);
        new HRESULT put_InputEventsAtOnce(int pinputEventsAtOnce);
        new HRESULT get_InputEventsAtOnce(out int pinputEventsAtOnce);
        new HRESULT put_maxEventCount(int pmaxEventCount);
        new HRESULT get_maxEventCount(out int pmaxEventCount);
        new HRESULT put_keepAliveInterval(int pkeepAliveInterval);
        new HRESULT get_keepAliveInterval(out int pkeepAliveInterval);
        new HRESULT put_shutdownTimeout(int pshutdownTimeout);
        new HRESULT get_shutdownTimeout(out int pshutdownTimeout);
        new HRESULT put_overallConnectionTimeout(int poverallConnectionTimeout);
        new HRESULT get_overallConnectionTimeout(out int poverallConnectionTimeout);
        new HRESULT put_singleConnectionTimeout(int psingleConnectionTimeout);
        new HRESULT get_singleConnectionTimeout(out int psingleConnectionTimeout);
        new HRESULT put_KeyboardType(int pkeyboardType);
        new HRESULT get_KeyboardType(out int pkeyboardType);
        new HRESULT put_KeyboardSubType(int pkeyboardSubType);
        new HRESULT get_KeyboardSubType(out int pkeyboardSubType);
        new HRESULT put_KeyboardFunctionKey(int pkeyboardFunctionKey);
        new HRESULT get_KeyboardFunctionKey(out int pkeyboardFunctionKey);
        new HRESULT put_WinceFixedPalette(int pwinceFixedPalette);
        new HRESULT get_WinceFixedPalette(out int pwinceFixedPalette);
        new HRESULT put_ConnectToServerConsole(VARIANT_BOOL pConnectToConsole);
        new HRESULT get_ConnectToServerConsole(out VARIANT_BOOL pConnectToConsole);
        new HRESULT put_BitmapPersistence(int pbitmapPersistence);
        new HRESULT get_BitmapPersistence(out int pbitmapPersistence);
        new HRESULT put_MinutesToIdleTimeout(int pminutesToIdleTimeout);
        new HRESULT get_MinutesToIdleTimeout(out int pminutesToIdleTimeout);
        new HRESULT put_SmartSizing(VARIANT_BOOL pfSmartSizing);
        new HRESULT get_SmartSizing(out VARIANT_BOOL pfSmartSizing);
        new HRESULT put_RdpdrLocalPrintingDocName(string pLocalPrintingDocName);
        new HRESULT get_RdpdrLocalPrintingDocName(out string pLocalPrintingDocName);
        new HRESULT put_RdpdrClipCleanTempDirString(string clipCleanTempDirString);
        new HRESULT get_RdpdrClipCleanTempDirString(out string clipCleanTempDirString);
        new HRESULT put_RdpdrClipPasteInfoString(string clipPasteInfoString);
        new HRESULT get_RdpdrClipPasteInfoString(out string clipPasteInfoString);
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_DisplayConnectionBar(VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT get_DisplayConnectionBar(out VARIANT_BOOL pDisplayConnectionBar);
        new HRESULT put_PinConnectionBar(VARIANT_BOOL pPinConnectionBar);
        new HRESULT get_PinConnectionBar(out VARIANT_BOOL pPinConnectionBar);
        new HRESULT put_GrabFocusOnConnect(VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT get_GrabFocusOnConnect(out VARIANT_BOOL pfGrabFocusOnConnect);
        new HRESULT put_LoadBalanceInfo(string pLBInfo);
        new HRESULT get_LoadBalanceInfo(out string pLBInfo);
        new HRESULT put_RedirectDrives(VARIANT_BOOL pRedirectDrives);
        new HRESULT get_RedirectDrives(out VARIANT_BOOL pRedirectDrives);
        new HRESULT put_RedirectPrinters(VARIANT_BOOL pRedirectPrinters);
        new HRESULT get_RedirectPrinters(out VARIANT_BOOL pRedirectPrinters);
        new HRESULT put_RedirectPorts(VARIANT_BOOL pRedirectPorts);
        new HRESULT get_RedirectPorts(out VARIANT_BOOL pRedirectPorts);
        new HRESULT put_RedirectSmartCards(VARIANT_BOOL pRedirectSmartCards);
        new HRESULT get_RedirectSmartCards(out VARIANT_BOOL pRedirectSmartCards);
        new HRESULT put_BitmapCache16BppSize(int pBitmapCache16BppSize);
        new HRESULT get_BitmapCache16BppSize(out int pBitmapCache16BppSize);
        new HRESULT put_BitmapCache24BppSize(int pBitmapCache24BppSize);
        new HRESULT get_BitmapCache24BppSize(out int pBitmapCache24BppSize);
        new HRESULT put_PerformanceFlags(int pDisableList);
        new HRESULT get_PerformanceFlags(out int pDisableList);
        // new HRESULT put_ConnectWithEndpoint(VARIANT _arg1);
        new HRESULT put_ConnectWithEndpoint(IntPtr _arg1);
        new HRESULT put_NotifyTSPublicKey(VARIANT_BOOL pfNotify);
        new HRESULT get_NotifyTSPublicKey(out VARIANT_BOOL pfNotify);

        new HRESULT get_CanAutoReconnect(out VARIANT_BOOL pfCanAutoReconnect);
        new HRESULT put_EnableAutoReconnect(VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT get_EnableAutoReconnect(out VARIANT_BOOL pfEnableAutoReconnect);
        new HRESULT put_MaxReconnectAttempts(int pMaxReconnectAttempts);
        new HRESULT get_MaxReconnectAttempts(out int pMaxReconnectAttempts);

        new HRESULT put_ConnectionBarShowMinimizeButton(VARIANT_BOOL pfShowMinimize);
        new HRESULT get_ConnectionBarShowMinimizeButton(out VARIANT_BOOL pfShowMinimize);
        new HRESULT put_ConnectionBarShowRestoreButton(VARIANT_BOOL pfShowRestore);
        new HRESULT get_ConnectionBarShowRestoreButton(out VARIANT_BOOL pfShowRestore);

        new HRESULT put_AuthenticationLevel(uint puiAuthLevel);
        new HRESULT get_AuthenticationLevel(out uint puiAuthLevel);

        new HRESULT put_RedirectClipboard(VARIANT_BOOL pfRedirectClipboard);
        new HRESULT get_RedirectClipboard(out VARIANT_BOOL pfRedirectClipboard);
        new HRESULT put_AudioRedirectionMode(uint puiAudioRedirectionMode);
        new HRESULT get_AudioRedirectionMode(out uint puiAudioRedirectionMode);
        new HRESULT put_ConnectionBarShowPinButton(VARIANT_BOOL pfShowPin);
        new HRESULT get_ConnectionBarShowPinButton(out VARIANT_BOOL pfShowPin);
        new HRESULT put_PublicMode(VARIANT_BOOL pfPublicMode);
        new HRESULT get_PublicMode(out VARIANT_BOOL pfPublicMode);
        new HRESULT put_RedirectDevices(VARIANT_BOOL pfRedirectPnPDevices);
        new HRESULT get_RedirectDevices(out VARIANT_BOOL pfRedirectPnPDevices);
        new HRESULT put_RedirectPOSDevices(VARIANT_BOOL pfRedirectPOSDevices);
        new HRESULT get_RedirectPOSDevices(out VARIANT_BOOL pfRedirectPOSDevices);
        new HRESULT put_BitmapVirtualCache32BppSize(int pBitmapVirtualCache32BppSize);
        new HRESULT get_BitmapVirtualCache32BppSize(out int pBitmapVirtualCache32BppSize);

        new HRESULT put_RelativeMouseMode(VARIANT_BOOL pfRelativeMouseMode);
        new HRESULT get_RelativeMouseMode(out VARIANT_BOOL pfRelativeMouseMode);
        new HRESULT get_AuthenticationServiceClass(out string pbstrAuthServiceClass);
        new HRESULT put_AuthenticationServiceClass(out string pbstrAuthServiceClass);
        new HRESULT get_PCB(out string bstrPCB);
        new HRESULT put_PCB(string bstrPCB);
        new HRESULT put_HotKeyFocusReleaseLeft(int HotKeyFocusReleaseLeft);
        new HRESULT get_HotKeyFocusReleaseLeft(out int HotKeyFocusReleaseLeft);
        new HRESULT put_HotKeyFocusReleaseRight(int HotKeyFocusReleaseRight);
        new HRESULT get_HotKeyFocusReleaseRight(out int HotKeyFocusReleaseRight);
        new HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        new HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        new HRESULT get_AuthenticationType(out uint puiAuthType);
        new HRESULT put_ConnectToAdministerServer(VARIANT_BOOL pConnectToAdministerServer);
        new HRESULT get_ConnectToAdministerServer(out VARIANT_BOOL pConnectToAdministerServer);

        new HRESULT put_AudioCaptureRedirectionMode(VARIANT_BOOL pfRedir);
        new HRESULT get_AudioCaptureRedirectionMode(out VARIANT_BOOL pfRedir);
        new HRESULT put_VideoPlaybackMode(uint pVideoPlaybackMode);
        new HRESULT get_VideoPlaybackMode(out uint pVideoPlaybackMode);
        new HRESULT put_EnableSuperPan(VARIANT_BOOL pfEnableSuperPan);
        new HRESULT get_EnableSuperPan(out VARIANT_BOOL pfEnableSuperPan);
        new HRESULT put_SuperPanAccelerationFactor(uint puAccelFactor);
        new HRESULT get_SuperPanAccelerationFactor(out uint puAccelFactor);
        new HRESULT put_NegotiateSecurityLayer(VARIANT_BOOL pfNegotiate);
        new HRESULT get_NegotiateSecurityLayer(out VARIANT_BOOL pfNegotiate);
        new HRESULT put_AudioQualityMode(uint pAudioQualityMode);
        new HRESULT get_AudioQualityMode(out uint pAudioQualityMode);
        new HRESULT put_RedirectDirectX(VARIANT_BOOL pfRedirectDirectX);
        new HRESULT get_RedirectDirectX(out VARIANT_BOOL pfRedirectDirectX);
        new HRESULT put_NetworkConnectionType(uint pConnectionType);
        new HRESULT get_NetworkConnectionType(out uint pConnectionType);

        HRESULT put_BandwidthDetection(VARIANT_BOOL pfAutodetect);
        HRESULT get_BandwidthDetection(out VARIANT_BOOL pfAutodetect);
        HRESULT put_ClientProtocolSpec(ClientSpec pClientMode);
        HRESULT get_ClientProtocolSpec(out ClientSpec pClientMode);
    }

    public enum ClientSpec
    {
        FullMode = 0,
        ThinClientMode = 1,
        SmallCacheMode = 2
    };


    [ComImport]
    [Guid("3d5b21ac-748d-41de-8f30-e15169586bd4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientTransportSettings3 : IMsRdpClientTransportSettings2
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_GatewayHostname(string pProxyHostname);
        new HRESULT get_GatewayHostname(out string pProxyHostname);
        new HRESULT put_GatewayUsageMethod(uint pulProxyUsageMethod);
        new HRESULT get_GatewayUsageMethod(out uint pulProxyUsageMethod);
        new HRESULT put_GatewayProfileUsageMethod(uint pulProxyProfileUsageMethod);
        new HRESULT get_GatewayProfileUsageMethod(out uint pulProxyProfileUsageMethod);
        new HRESULT put_GatewayCredsSource(uint pulProxyCredsSource);
        new HRESULT get_GatewayCredsSource(out uint pulProxyCredsSource);
        new HRESULT put_GatewayUserSelectedCredsSource(uint pulProxyCredsSource);
        new HRESULT get_GatewayUserSelectedCredsSource(out uint pulProxyCredsSource);
        new HRESULT get_GatewayIsSupported(out int pfProxyIsSupported);
        new HRESULT get_GatewayDefaultUsageMethod(out uint pulProxyDefaultUsageMethod);

        new HRESULT put_GatewayCredSharing(uint pulProxyCredSharing);
        new HRESULT get_GatewayCredSharing(out uint pulProxyCredSharing);
        new HRESULT put_GatewayPreAuthRequirement(uint pulProxyPreAuthRequirement);
        new HRESULT get_GatewayPreAuthRequirement(out uint pulProxyPreAuthRequirement);
        new HRESULT put_GatewayPreAuthServerAddr(string pstringProxyPreAuthServerAddr);
        new HRESULT get_GatewayPreAuthServerAddr(out string pstringProxyPreAuthServerAddr);
        new HRESULT put_GatewaySupportUrl(string pstringProxySupportUrl);
        new HRESULT get_GatewaySupportUrl(out string pstringProxySupportUrl);
        new HRESULT put_GatewayEncryptedOtpCookie(string pstringEncryptedOtpCookie);
        new HRESULT get_GatewayEncryptedOtpCookie(out string pstringEncryptedOtpCookie);
        new HRESULT put_GatewayEncryptedOtpCookieSize(uint pulEncryptedOtpCookieSize);
        new HRESULT get_GatewayEncryptedOtpCookieSize(out uint pulEncryptedOtpCookieSize);
        new HRESULT put_GatewayUsername(string pProxyUsername);
        new HRESULT get_GatewayUsername(out string pProxyUsername);
        new HRESULT put_GatewayDomain(string pProxyDomain);
        new HRESULT get_GatewayDomain(out string pProxyDomain);
        new HRESULT put_GatewayPassword(string _arg1);

        HRESULT put_GatewayCredSourceCookie(uint pulProxyCredSourceCookie);
        HRESULT get_GatewayCredSourceCookie(out uint pulProxyCredSourceCookie);
        HRESULT put_GatewayAuthCookieServerAddr(string pstringProxyAuthCookieServerAddr);
        HRESULT get_GatewayAuthCookieServerAddr(out string pstringProxyAuthCookieServerAddr);
        HRESULT put_GatewayEncryptedAuthCookie(string pstringEncryptedAuthCookie);
        HRESULT get_GatewayEncryptedAuthCookie(out string pstringEncryptedAuthCookie);
        HRESULT put_GatewayEncryptedAuthCookieSize(uint pulEncryptedAuthCookieSize);
        HRESULT get_GatewayEncryptedAuthCookieSize(out uint pulEncryptedAuthCookieSize);
        HRESULT put_GatewayAuthLoginPage(string pstringProxyAuthLoginPage);
        HRESULT get_GatewayAuthLoginPage(out string pstringProxyAuthLoginPage);
    }



    [ComImport]
    [Guid("011c3236-4d81-4515-9143-067ab630d299")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientTransportSettings4 : IMsRdpClientTransportSettings3
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_GatewayHostname(string pProxyHostname);
        new HRESULT get_GatewayHostname(out string pProxyHostname);
        new HRESULT put_GatewayUsageMethod(uint pulProxyUsageMethod);
        new HRESULT get_GatewayUsageMethod(out uint pulProxyUsageMethod);
        new HRESULT put_GatewayProfileUsageMethod(uint pulProxyProfileUsageMethod);
        new HRESULT get_GatewayProfileUsageMethod(out uint pulProxyProfileUsageMethod);
        new HRESULT put_GatewayCredsSource(uint pulProxyCredsSource);
        new HRESULT get_GatewayCredsSource(out uint pulProxyCredsSource);
        new HRESULT put_GatewayUserSelectedCredsSource(uint pulProxyCredsSource);
        new HRESULT get_GatewayUserSelectedCredsSource(out uint pulProxyCredsSource);
        new HRESULT get_GatewayIsSupported(out int pfProxyIsSupported);
        new HRESULT get_GatewayDefaultUsageMethod(out uint pulProxyDefaultUsageMethod);

        new HRESULT put_GatewayCredSharing(uint pulProxyCredSharing);
        new HRESULT get_GatewayCredSharing(out uint pulProxyCredSharing);
        new HRESULT put_GatewayPreAuthRequirement(uint pulProxyPreAuthRequirement);
        new HRESULT get_GatewayPreAuthRequirement(out uint pulProxyPreAuthRequirement);
        new HRESULT put_GatewayPreAuthServerAddr(string pstringProxyPreAuthServerAddr);
        new HRESULT get_GatewayPreAuthServerAddr(out string pstringProxyPreAuthServerAddr);
        new HRESULT put_GatewaySupportUrl(string pstringProxySupportUrl);
        new HRESULT get_GatewaySupportUrl(out string pstringProxySupportUrl);
        new HRESULT put_GatewayEncryptedOtpCookie(string pstringEncryptedOtpCookie);
        new HRESULT get_GatewayEncryptedOtpCookie(out string pstringEncryptedOtpCookie);
        new HRESULT put_GatewayEncryptedOtpCookieSize(uint pulEncryptedOtpCookieSize);
        new HRESULT get_GatewayEncryptedOtpCookieSize(out uint pulEncryptedOtpCookieSize);
        new HRESULT put_GatewayUsername(string pProxyUsername);
        new HRESULT get_GatewayUsername(out string pProxyUsername);
        new HRESULT put_GatewayDomain(string pProxyDomain);
        new HRESULT get_GatewayDomain(out string pProxyDomain);
        new HRESULT put_GatewayPassword(string _arg1);

        new HRESULT put_GatewayCredSourceCookie(uint pulProxyCredSourceCookie);
        new HRESULT get_GatewayCredSourceCookie(out uint pulProxyCredSourceCookie);
        new HRESULT put_GatewayAuthCookieServerAddr(string pstringProxyAuthCookieServerAddr);
        new HRESULT get_GatewayAuthCookieServerAddr(out string pstringProxyAuthCookieServerAddr);
        new HRESULT put_GatewayEncryptedAuthCookie(string pstringEncryptedAuthCookie);
        new HRESULT get_GatewayEncryptedAuthCookie(out string pstringEncryptedAuthCookie);
        new HRESULT put_GatewayEncryptedAuthCookieSize(uint pulEncryptedAuthCookieSize);
        new HRESULT get_GatewayEncryptedAuthCookieSize(out uint pulEncryptedAuthCookieSize);
        new HRESULT put_GatewayAuthLoginPage(string pstringProxyAuthLoginPage);
        new HRESULT get_GatewayAuthLoginPage(out string pstringProxyAuthLoginPage);

        HRESULT put_GatewayBrokeringType(uint _arg1);
    }



    [ComImport]
    [Guid("25f2ce20-8b1d-4971-a7cd-549dae201fc0")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientSecuredSettings2 : IMsRdpClientSecuredSettings
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_StartProgram(string pStartProgram);
        new HRESULT get_StartProgram(out string pStartProgram);
        new HRESULT put_WorkDir(string pWorkDir);
        new HRESULT get_WorkDir(out string pWorkDir);
        new HRESULT put_FullScreen(int pfFullScreen);
        new HRESULT get_FullScreen(out int pfFullScreen);

        new HRESULT put_KeyboardHookMode(int pkeyboardHookMode);
        new HRESULT get_KeyboardHookMode(out int pkeyboardHookMode);
        new HRESULT put_AudioRedirectionMode(int pAudioRedirectionMode);
        new HRESULT get_AudioRedirectionMode(out int pAudioRedirectionMode);

        HRESULT get_PCB(out string bstrPCB);
        HRESULT put_PCB(string bstrPCB);
    }

    [ComImport]
    [Guid("92c38a7d-241a-418c-9936-099872c9af20")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITSRemoteProgram2 : ITSRemoteProgram
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_RemoteProgramMode(VARIANT_BOOL pvboolRemoteProgramMode);
        new HRESULT get_RemoteProgramMode(out VARIANT_BOOL pvboolRemoteProgramMode);
        [PreserveSig]
        new HRESULT ServerStartProgram(string bstrExecutablePath, string bstrFilePath, string bstrWorkingDirectory,
            VARIANT_BOOL vbExpandEnvVarInWorkingDirectoryOnServer, string bstrArguments, VARIANT_BOOL vbExpandEnvVarInArgumentsOnServer);

        HRESULT put_RemoteApplicationName(string _arg1);
        HRESULT put_RemoteApplicationProgram(string _arg1);
        HRESULT put_RemoteApplicationArgs(string _arg1);
    }

    [ComImport]
    [Guid("4b84ea77-acea-418c-881a-4a8c28ab1510")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITSRemoteProgram3 : ITSRemoteProgram2
    {
        #region <IDispatch>
        new int GetTypeInfoCount();
        [return: MarshalAs(UnmanagedType.Interface)]
        new IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
        [PreserveSig]
        new HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
            [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        new HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
            [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        #endregion

        new HRESULT put_RemoteProgramMode(VARIANT_BOOL pvboolRemoteProgramMode);
        new HRESULT get_RemoteProgramMode(out VARIANT_BOOL pvboolRemoteProgramMode);
        [PreserveSig]
        new HRESULT ServerStartProgram(string bstrExecutablePath, string bstrFilePath, string bstrWorkingDirectory,
            VARIANT_BOOL vbExpandEnvVarInWorkingDirectoryOnServer, string bstrArguments, VARIANT_BOOL vbExpandEnvVarInArgumentsOnServer);

        new HRESULT put_RemoteApplicationName(string _arg1);
        new HRESULT put_RemoteApplicationProgram(string _arg1);
        new HRESULT put_RemoteApplicationArgs(string _arg1);

        HRESULT ServerStartApp(string bstrAppUserModelId, string bstrArguments, VARIANT_BOOL vbExpandEnvVarInArgumentsOnServer);
    }

    [ComImport]
    [Guid("c1e6743a-41c1-4a74-832a-0dd06c1c7a0e")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsTscNonScriptable
    {
        HRESULT put_ClearTextPassword(string _arg1);
        HRESULT put_PortablePassword(string pPortablePass);
        HRESULT get_PortablePassword(out string pPortablePass);
        HRESULT put_PortableSalt(string pPortableSalt);
        HRESULT get_PortableSalt(out string pPortableSalt);
        HRESULT put_BinaryPassword(string pBinaryPassword);
        HRESULT get_BinaryPassword(out string pBinaryPassword);
        HRESULT put_BinarySalt(string pSalt);
        HRESULT get_BinarySalt(out string pSalt);
        HRESULT ResetPassword();
    }

    [ComImport]
    [Guid("2f079c4c-87b2-4afd-97ab-20cdb43038ae")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientNonScriptable : IMsTscNonScriptable
    {
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_PortablePassword(string pPortablePass);
        new HRESULT get_PortablePassword(out string pPortablePass);
        new HRESULT put_PortableSalt(string pPortableSalt);
        new HRESULT get_PortableSalt(out string pPortableSalt);
        new HRESULT put_BinaryPassword(string pBinaryPassword);
        new HRESULT get_BinaryPassword(out string pBinaryPassword);
        new HRESULT put_BinarySalt(string pSalt);
        new HRESULT get_BinarySalt(out string pSalt);
        new HRESULT ResetPassword();

        HRESULT NotifyRedirectDeviceChange(uint wParam, IntPtr lParam);
        HRESULT SendKeys(int numKeys, VARIANT_BOOL pbArrayKeyUp, int plKeyData);
    }

    [ComImport]
    [Guid("17a5e535-4072-4fa4-af32-c8d0d47345e9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientNonScriptable2 : IMsRdpClientNonScriptable
    {
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_PortablePassword(string pPortablePass);
        new HRESULT get_PortablePassword(out string pPortablePass);
        new HRESULT put_PortableSalt(string pPortableSalt);
        new HRESULT get_PortableSalt(out string pPortableSalt);
        new HRESULT put_BinaryPassword(string pBinaryPassword);
        new HRESULT get_BinaryPassword(out string pBinaryPassword);
        new HRESULT put_BinarySalt(string pSalt);
        new HRESULT get_BinarySalt(out string pSalt);
        new HRESULT ResetPassword();

        new HRESULT NotifyRedirectDeviceChange(uint wParam, IntPtr lParam);
        new HRESULT SendKeys(int numKeys, VARIANT_BOOL pbArrayKeyUp, int plKeyData);

        HRESULT put_UIParentWindowHandle(IntPtr phwndUIParentWindowHandle);
        HRESULT  get_UIParentWindowHandle(out IntPtr phwndUIParentWindowHandle);
    }

    [ComImport]
    [Guid("b3378d90-0728-45c7-8ed7-b6159fb92219")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientNonScriptable3 : IMsRdpClientNonScriptable2
    {
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_PortablePassword(string pPortablePass);
        new HRESULT get_PortablePassword(out string pPortablePass);
        new HRESULT put_PortableSalt(string pPortableSalt);
        new HRESULT get_PortableSalt(out string pPortableSalt);
        new HRESULT put_BinaryPassword(string pBinaryPassword);
        new HRESULT get_BinaryPassword(out string pBinaryPassword);
        new HRESULT put_BinarySalt(string pSalt);
        new HRESULT get_BinarySalt(out string pSalt);
        new HRESULT ResetPassword();

        new HRESULT NotifyRedirectDeviceChange(uint wParam, IntPtr lParam);
        new HRESULT SendKeys(int numKeys, VARIANT_BOOL pbArrayKeyUp, int plKeyData);

        new HRESULT put_UIParentWindowHandle(IntPtr phwndUIParentWindowHandle);
        new HRESULT get_UIParentWindowHandle(out IntPtr phwndUIParentWindowHandle);

        HRESULT put_ShowRedirectionWarningDialog(VARIANT_BOOL pfShowRdrDlg);
        HRESULT get_ShowRedirectionWarningDialog(out VARIANT_BOOL pfShowRdrDlg);
        HRESULT put_PromptForCredentials(VARIANT_BOOL pfPrompt);
        HRESULT get_PromptForCredentials(out VARIANT_BOOL pfPrompt);
        HRESULT put_NegotiateSecurityLayer(VARIANT_BOOL pfNegotiate);
        HRESULT get_NegotiateSecurityLayer(out VARIANT_BOOL pfNegotiate);
        HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        HRESULT put_RedirectDynamicDrives(VARIANT_BOOL pfRedirectDynamicDrives);
        HRESULT get_RedirectDynamicDrives(out VARIANT_BOOL pfRedirectDynamicDrives);
        HRESULT put_RedirectDynamicDevices(VARIANT_BOOL pfRedirectDynamicDevices);
        HRESULT get_RedirectDynamicDevices(out VARIANT_BOOL pfRedirectDynamicDevices);
        HRESULT get_DeviceCollection(out IMsRdpDeviceCollection ppDeviceCollection);
        HRESULT get_DriveCollection(out IMsRdpDriveCollection ppDeviceCollection);
        HRESULT put_WarnAboutSendingCredentials(VARIANT_BOOL pfWarn);
        HRESULT get_WarnAboutSendingCredentials(out VARIANT_BOOL pfWarn);
        HRESULT put_WarnAboutClipboardRedirection(VARIANT_BOOL pfWarn);
        HRESULT get_WarnAboutClipboardRedirection(out VARIANT_BOOL pfWarn);
        HRESULT put_ConnectionBarText(string pConnectionBarText);
        HRESULT get_ConnectionBarText(out string pConnectionBarText);
    }

    [ComImport]
    [Guid("f50fa8aa-1c7d-4f59-b15c-a90cacae1fcb")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientNonScriptable4 : IMsRdpClientNonScriptable3
    {
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_PortablePassword(string pPortablePass);
        new HRESULT get_PortablePassword(out string pPortablePass);
        new HRESULT put_PortableSalt(string pPortableSalt);
        new HRESULT get_PortableSalt(out string pPortableSalt);
        new HRESULT put_BinaryPassword(string pBinaryPassword);
        new HRESULT get_BinaryPassword(out string pBinaryPassword);
        new HRESULT put_BinarySalt(string pSalt);
        new HRESULT get_BinarySalt(out string pSalt);
        new HRESULT ResetPassword();

        new HRESULT NotifyRedirectDeviceChange(uint wParam, IntPtr lParam);
        new HRESULT SendKeys(int numKeys, VARIANT_BOOL pbArrayKeyUp, int plKeyData);

        new HRESULT put_UIParentWindowHandle(IntPtr phwndUIParentWindowHandle);
        new HRESULT get_UIParentWindowHandle(out IntPtr phwndUIParentWindowHandle);

        new HRESULT put_ShowRedirectionWarningDialog(VARIANT_BOOL pfShowRdrDlg);
        new HRESULT get_ShowRedirectionWarningDialog(out VARIANT_BOOL pfShowRdrDlg);
        new HRESULT put_PromptForCredentials(VARIANT_BOOL pfPrompt);
        new HRESULT get_PromptForCredentials(out VARIANT_BOOL pfPrompt);
        new HRESULT put_NegotiateSecurityLayer(VARIANT_BOOL pfNegotiate);
        new HRESULT get_NegotiateSecurityLayer(out VARIANT_BOOL pfNegotiate);
        new HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        new HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        new HRESULT put_RedirectDynamicDrives(VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT get_RedirectDynamicDrives(out VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT put_RedirectDynamicDevices(VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_RedirectDynamicDevices(out VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_DeviceCollection(out IMsRdpDeviceCollection ppDeviceCollection);
        new HRESULT get_DriveCollection(out IMsRdpDriveCollection ppDeviceCollection);
        new HRESULT put_WarnAboutSendingCredentials(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutSendingCredentials(out VARIANT_BOOL pfWarn);
        new HRESULT put_WarnAboutClipboardRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutClipboardRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_ConnectionBarText(string pConnectionBarText);
        new HRESULT get_ConnectionBarText(out string pConnectionBarText);

        HRESULT put_RedirectionWarningType(RedirectionWarningType pWrnType);
        HRESULT get_RedirectionWarningType(out RedirectionWarningType pWrnType);
        HRESULT put_MarkRdpSettingsSecure(VARIANT_BOOL pfRdpSecure);
        HRESULT get_MarkRdpSettingsSecure(out VARIANT_BOOL pfRdpSecure);
        //HRESULT put_PublisherCertificateChain(VARIANT pVarCert);
        HRESULT put_PublisherCertificateChain(IntPtr pVarCert);
        //HRESULT get_PublisherCertificateChain(out VARIANT pVarCert);
        HRESULT get_PublisherCertificateChain(out IntPtr pVarCert);
        HRESULT put_WarnAboutPrinterRedirection(VARIANT_BOOL pfWarn);
        HRESULT get_WarnAboutPrinterRedirection(out VARIANT_BOOL pfWarn);
        HRESULT put_AllowCredentialSaving(VARIANT_BOOL pfAllowSave);
        HRESULT get_AllowCredentialSaving(out VARIANT_BOOL pfAllowSave);
        HRESULT put_PromptForCredsOnClient(VARIANT_BOOL pfPromptForCredsOnClient);
        HRESULT get_PromptForCredsOnClient(out VARIANT_BOOL pfPromptForCredsOnClient);
        HRESULT put_LaunchedViaClientShellInterface(VARIANT_BOOL pfLaunchedViaClientShellInterface);
        HRESULT get_LaunchedViaClientShellInterface(out VARIANT_BOOL pfLaunchedViaClientShellInterface);
        HRESULT put_TrustedZoneSite(VARIANT_BOOL pfIsTrustedZone);
        HRESULT get_TrustedZoneSite(out VARIANT_BOOL pfIsTrustedZone);
    }

    public enum RedirectionWarningType
    {
        RedirectionWarningTypeDefault = 0,
        RedirectionWarningTypeUnsigned = 1,
        RedirectionWarningTypeUnknown = 2,
        RedirectionWarningTypeUser = 3,
        RedirectionWarningTypeThirdPartySigned = 4,
        RedirectionWarningTypeTrusted = 5,
        RedirectionWarningTypeMax = 5
    };

    [ComImport]
    [Guid("4f6996d5-d7b1-412c-b0ff-063718566907")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientNonScriptable5 : IMsRdpClientNonScriptable4
    {
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_PortablePassword(string pPortablePass);
        new HRESULT get_PortablePassword(out string pPortablePass);
        new HRESULT put_PortableSalt(string pPortableSalt);
        new HRESULT get_PortableSalt(out string pPortableSalt);
        new HRESULT put_BinaryPassword(string pBinaryPassword);
        new HRESULT get_BinaryPassword(out string pBinaryPassword);
        new HRESULT put_BinarySalt(string pSalt);
        new HRESULT get_BinarySalt(out string pSalt);
        new HRESULT ResetPassword();

        new HRESULT NotifyRedirectDeviceChange(uint wParam, IntPtr lParam);
        new HRESULT SendKeys(int numKeys, VARIANT_BOOL pbArrayKeyUp, int plKeyData);

        new HRESULT put_UIParentWindowHandle(IntPtr phwndUIParentWindowHandle);
        new HRESULT get_UIParentWindowHandle(out IntPtr phwndUIParentWindowHandle);

        new HRESULT put_ShowRedirectionWarningDialog(VARIANT_BOOL pfShowRdrDlg);
        new HRESULT get_ShowRedirectionWarningDialog(out VARIANT_BOOL pfShowRdrDlg);
        new HRESULT put_PromptForCredentials(VARIANT_BOOL pfPrompt);
        new HRESULT get_PromptForCredentials(out VARIANT_BOOL pfPrompt);
        new HRESULT put_NegotiateSecurityLayer(VARIANT_BOOL pfNegotiate);
        new HRESULT get_NegotiateSecurityLayer(out VARIANT_BOOL pfNegotiate);
        new HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        new HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        new HRESULT put_RedirectDynamicDrives(VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT get_RedirectDynamicDrives(out VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT put_RedirectDynamicDevices(VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_RedirectDynamicDevices(out VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_DeviceCollection(out IMsRdpDeviceCollection ppDeviceCollection);
        new HRESULT get_DriveCollection(out IMsRdpDriveCollection ppDeviceCollection);
        new HRESULT put_WarnAboutSendingCredentials(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutSendingCredentials(out VARIANT_BOOL pfWarn);
        new HRESULT put_WarnAboutClipboardRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutClipboardRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_ConnectionBarText(string pConnectionBarText);
        new HRESULT get_ConnectionBarText(out string pConnectionBarText);

        new HRESULT put_RedirectionWarningType(RedirectionWarningType pWrnType);
        new HRESULT get_RedirectionWarningType(out RedirectionWarningType pWrnType);
        new HRESULT put_MarkRdpSettingsSecure(VARIANT_BOOL pfRdpSecure);
        new HRESULT get_MarkRdpSettingsSecure(out VARIANT_BOOL pfRdpSecure);
        //new HRESULT put_PublisherCertificateChain(VARIANT pVarCert);
        new HRESULT put_PublisherCertificateChain(IntPtr pVarCert);
        //new HRESULT get_PublisherCertificateChain(out VARIANT pVarCert);
        new HRESULT get_PublisherCertificateChain(out IntPtr pVarCert);
        new HRESULT put_WarnAboutPrinterRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutPrinterRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_AllowCredentialSaving(VARIANT_BOOL pfAllowSave);
        new HRESULT get_AllowCredentialSaving(out VARIANT_BOOL pfAllowSave);
        new HRESULT put_PromptForCredsOnClient(VARIANT_BOOL pfPromptForCredsOnClient);
        new HRESULT get_PromptForCredsOnClient(out VARIANT_BOOL pfPromptForCredsOnClient);
        new HRESULT put_LaunchedViaClientShellInterface(VARIANT_BOOL pfLaunchedViaClientShellInterface);
        new HRESULT get_LaunchedViaClientShellInterface(out VARIANT_BOOL pfLaunchedViaClientShellInterface);
        new HRESULT put_TrustedZoneSite(VARIANT_BOOL pfIsTrustedZone);
        new HRESULT get_TrustedZoneSite(out VARIANT_BOOL pfIsTrustedZone);

        HRESULT put_UseMultimon(VARIANT_BOOL pfUseMultimon);
        HRESULT get_UseMultimon(out VARIANT_BOOL pfUseMultimon);
        HRESULT get_RemoteMonitorCount(out uint pcRemoteMonitors);
        HRESULT GetRemoteMonitorsBoundingBox(out int pLeft, out int pTop, out int pRight, out int pBottom);
        HRESULT get_RemoteMonitorLayoutMatchesLocal(out VARIANT_BOOL pfRemoteMatchesLocal);
        HRESULT put_DisableConnectionBar(VARIANT_BOOL _arg1);
        HRESULT put_DisableRemoteAppCapsCheck(VARIANT_BOOL pfDisableRemoteAppCapsCheck);
        HRESULT get_DisableRemoteAppCapsCheck(out VARIANT_BOOL pfDisableRemoteAppCapsCheck);
        HRESULT put_WarnAboutDirectXRedirection(VARIANT_BOOL pfWarn);
        HRESULT get_WarnAboutDirectXRedirection(out VARIANT_BOOL pfWarn);
        HRESULT put_AllowPromptingForCredentials(VARIANT_BOOL pfAllow);
        HRESULT get_AllowPromptingForCredentials(out VARIANT_BOOL pfAllow);
    }

    [ComImport]
    [Guid("05293249-b28b-4bd8-be64-1b2f496b910e")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientNonScriptable6 : IMsRdpClientNonScriptable5
    {
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_PortablePassword(string pPortablePass);
        new HRESULT get_PortablePassword(out string pPortablePass);
        new HRESULT put_PortableSalt(string pPortableSalt);
        new HRESULT get_PortableSalt(out string pPortableSalt);
        new HRESULT put_BinaryPassword(string pBinaryPassword);
        new HRESULT get_BinaryPassword(out string pBinaryPassword);
        new HRESULT put_BinarySalt(string pSalt);
        new HRESULT get_BinarySalt(out string pSalt);
        new HRESULT ResetPassword();

        new HRESULT NotifyRedirectDeviceChange(uint wParam, IntPtr lParam);
        new HRESULT SendKeys(int numKeys, VARIANT_BOOL pbArrayKeyUp, int plKeyData);

        new HRESULT put_UIParentWindowHandle(IntPtr phwndUIParentWindowHandle);
        new HRESULT get_UIParentWindowHandle(out IntPtr phwndUIParentWindowHandle);

        new HRESULT put_ShowRedirectionWarningDialog(VARIANT_BOOL pfShowRdrDlg);
        new HRESULT get_ShowRedirectionWarningDialog(out VARIANT_BOOL pfShowRdrDlg);
        new HRESULT put_PromptForCredentials(VARIANT_BOOL pfPrompt);
        new HRESULT get_PromptForCredentials(out VARIANT_BOOL pfPrompt);
        new HRESULT put_NegotiateSecurityLayer(VARIANT_BOOL pfNegotiate);
        new HRESULT get_NegotiateSecurityLayer(out VARIANT_BOOL pfNegotiate);
        new HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        new HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        new HRESULT put_RedirectDynamicDrives(VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT get_RedirectDynamicDrives(out VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT put_RedirectDynamicDevices(VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_RedirectDynamicDevices(out VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_DeviceCollection(out IMsRdpDeviceCollection ppDeviceCollection);
        new HRESULT get_DriveCollection(out IMsRdpDriveCollection ppDeviceCollection);
        new HRESULT put_WarnAboutSendingCredentials(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutSendingCredentials(out VARIANT_BOOL pfWarn);
        new HRESULT put_WarnAboutClipboardRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutClipboardRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_ConnectionBarText(string pConnectionBarText);
        new HRESULT get_ConnectionBarText(out string pConnectionBarText);

        new HRESULT put_RedirectionWarningType(RedirectionWarningType pWrnType);
        new HRESULT get_RedirectionWarningType(out RedirectionWarningType pWrnType);
        new HRESULT put_MarkRdpSettingsSecure(VARIANT_BOOL pfRdpSecure);
        new HRESULT get_MarkRdpSettingsSecure(out VARIANT_BOOL pfRdpSecure);
        //new HRESULT put_PublisherCertificateChain(VARIANT pVarCert);
        new HRESULT put_PublisherCertificateChain(IntPtr pVarCert);
        //new HRESULT get_PublisherCertificateChain(out VARIANT pVarCert);
        new HRESULT get_PublisherCertificateChain(out IntPtr pVarCert);
        new HRESULT put_WarnAboutPrinterRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutPrinterRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_AllowCredentialSaving(VARIANT_BOOL pfAllowSave);
        new HRESULT get_AllowCredentialSaving(out VARIANT_BOOL pfAllowSave);
        new HRESULT put_PromptForCredsOnClient(VARIANT_BOOL pfPromptForCredsOnClient);
        new HRESULT get_PromptForCredsOnClient(out VARIANT_BOOL pfPromptForCredsOnClient);
        new HRESULT put_LaunchedViaClientShellInterface(VARIANT_BOOL pfLaunchedViaClientShellInterface);
        new HRESULT get_LaunchedViaClientShellInterface(out VARIANT_BOOL pfLaunchedViaClientShellInterface);
        new HRESULT put_TrustedZoneSite(VARIANT_BOOL pfIsTrustedZone);
        new HRESULT get_TrustedZoneSite(out VARIANT_BOOL pfIsTrustedZone);

        new HRESULT put_UseMultimon(VARIANT_BOOL pfUseMultimon);
        new HRESULT get_UseMultimon(out VARIANT_BOOL pfUseMultimon);
        new HRESULT get_RemoteMonitorCount(out uint pcRemoteMonitors);
        new HRESULT GetRemoteMonitorsBoundingBox(out int pLeft, out int pTop, out int pRight, out int pBottom);
        new HRESULT get_RemoteMonitorLayoutMatchesLocal(out VARIANT_BOOL pfRemoteMatchesLocal);
        new HRESULT put_DisableConnectionBar(VARIANT_BOOL _arg1);
        new HRESULT put_DisableRemoteAppCapsCheck(VARIANT_BOOL pfDisableRemoteAppCapsCheck);
        new HRESULT get_DisableRemoteAppCapsCheck(out VARIANT_BOOL pfDisableRemoteAppCapsCheck);
        new HRESULT put_WarnAboutDirectXRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutDirectXRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_AllowPromptingForCredentials(VARIANT_BOOL pfAllow);
        new HRESULT get_AllowPromptingForCredentials(out VARIANT_BOOL pfAllow);

        HRESULT SendLocation2D(double latitude, double longitude);
        HRESULT SendLocation3D(double latitude, double longitude, int altitude);
    }

    [ComImport]
    [Guid("56540617-d281-488c-8738-6a8fdf64a118")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpDeviceCollection
    {
        HRESULT RescanDevices(VARIANT_BOOL vboolDynRedir);
        HRESULT get_DeviceByIndex(uint index, out IMsRdpDevice ppDevice);
        HRESULT get_DeviceById(string devInstanceId, out IMsRdpDevice ppDevice);
        HRESULT get_DeviceCount(out uint pDeviceCount);
    }

    [ComImport]
    [Guid("60c3b9c8-9e92-4f5e-a3e7-604a912093ea")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpDevice
    {
        HRESULT get_DeviceInstanceId(out string pDevInstanceId);
        HRESULT get_FriendlyName(out string pFriendlyName);
        HRESULT get_DeviceDescription(out string pDeviceDescription);
        HRESULT put_RedirectionState(VARIANT_BOOL pvboolRedirState);
        HRESULT get_RedirectionState(out VARIANT_BOOL pvboolRedirState);
    }

    [ComImport]
    [Guid("7ff17599-da2c-4677-ad35-f60c04fe1585")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpDriveCollection
    {
        HRESULT RescanDrives(VARIANT_BOOL vboolDynRedir);
        HRESULT get_DriveByIndex(uint index, out IMsRdpDrive ppDevice);
        HRESULT get_DriveCount(out uint pDriveCount);
    }

    [ComImport]
    [Guid("d28b5458-f694-47a8-8e61-40356a767e46")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpDrive
    {
        HRESULT get_Name(out string pName);
        HRESULT put_RedirectionState(VARIANT_BOOL pvboolRedirState);
        HRESULT get_RedirectionState(out VARIANT_BOOL pvboolRedirState);
    }

    [ComImport]
    [Guid("71b4a60a-fe21-46d8-a39b-8e32ba0c5ecc")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClientNonScriptable7 : IMsRdpClientNonScriptable6
    {
        new HRESULT put_ClearTextPassword(string _arg1);
        new HRESULT put_PortablePassword(string pPortablePass);
        new HRESULT get_PortablePassword(out string pPortablePass);
        new HRESULT put_PortableSalt(string pPortableSalt);
        new HRESULT get_PortableSalt(out string pPortableSalt);
        new HRESULT put_BinaryPassword(string pBinaryPassword);
        new HRESULT get_BinaryPassword(out string pBinaryPassword);
        new HRESULT put_BinarySalt(string pSalt);
        new HRESULT get_BinarySalt(out string pSalt);
        new HRESULT ResetPassword();

        new HRESULT NotifyRedirectDeviceChange(uint wParam, IntPtr lParam);
        new HRESULT SendKeys(int numKeys, VARIANT_BOOL pbArrayKeyUp, int plKeyData);

        new HRESULT put_UIParentWindowHandle(IntPtr phwndUIParentWindowHandle);
        new HRESULT get_UIParentWindowHandle(out IntPtr phwndUIParentWindowHandle);

        new HRESULT put_ShowRedirectionWarningDialog(VARIANT_BOOL pfShowRdrDlg);
        new HRESULT get_ShowRedirectionWarningDialog(out VARIANT_BOOL pfShowRdrDlg);
        new HRESULT put_PromptForCredentials(VARIANT_BOOL pfPrompt);
        new HRESULT get_PromptForCredentials(out VARIANT_BOOL pfPrompt);
        new HRESULT put_NegotiateSecurityLayer(VARIANT_BOOL pfNegotiate);
        new HRESULT get_NegotiateSecurityLayer(out VARIANT_BOOL pfNegotiate);
        new HRESULT put_EnableCredSspSupport(VARIANT_BOOL pfEnableSupport);
        new HRESULT get_EnableCredSspSupport(out VARIANT_BOOL pfEnableSupport);
        new HRESULT put_RedirectDynamicDrives(VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT get_RedirectDynamicDrives(out VARIANT_BOOL pfRedirectDynamicDrives);
        new HRESULT put_RedirectDynamicDevices(VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_RedirectDynamicDevices(out VARIANT_BOOL pfRedirectDynamicDevices);
        new HRESULT get_DeviceCollection(out IMsRdpDeviceCollection ppDeviceCollection);
        new HRESULT get_DriveCollection(out IMsRdpDriveCollection ppDeviceCollection);
        new HRESULT put_WarnAboutSendingCredentials(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutSendingCredentials(out VARIANT_BOOL pfWarn);
        new HRESULT put_WarnAboutClipboardRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutClipboardRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_ConnectionBarText(string pConnectionBarText);
        new HRESULT get_ConnectionBarText(out string pConnectionBarText);

        new HRESULT put_RedirectionWarningType(RedirectionWarningType pWrnType);
        new HRESULT get_RedirectionWarningType(out RedirectionWarningType pWrnType);
        new HRESULT put_MarkRdpSettingsSecure(VARIANT_BOOL pfRdpSecure);
        new HRESULT get_MarkRdpSettingsSecure(out VARIANT_BOOL pfRdpSecure);
        //new HRESULT put_PublisherCertificateChain(VARIANT pVarCert);
        new HRESULT put_PublisherCertificateChain(IntPtr pVarCert);
        //new HRESULT get_PublisherCertificateChain(out VARIANT pVarCert);
        new HRESULT get_PublisherCertificateChain(out IntPtr pVarCert);
        new HRESULT put_WarnAboutPrinterRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutPrinterRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_AllowCredentialSaving(VARIANT_BOOL pfAllowSave);
        new HRESULT get_AllowCredentialSaving(out VARIANT_BOOL pfAllowSave);
        new HRESULT put_PromptForCredsOnClient(VARIANT_BOOL pfPromptForCredsOnClient);
        new HRESULT get_PromptForCredsOnClient(out VARIANT_BOOL pfPromptForCredsOnClient);
        new HRESULT put_LaunchedViaClientShellInterface(VARIANT_BOOL pfLaunchedViaClientShellInterface);
        new HRESULT get_LaunchedViaClientShellInterface(out VARIANT_BOOL pfLaunchedViaClientShellInterface);
        new HRESULT put_TrustedZoneSite(VARIANT_BOOL pfIsTrustedZone);
        new HRESULT get_TrustedZoneSite(out VARIANT_BOOL pfIsTrustedZone);

        new HRESULT put_UseMultimon(VARIANT_BOOL pfUseMultimon);
        new HRESULT get_UseMultimon(out VARIANT_BOOL pfUseMultimon);
        new HRESULT get_RemoteMonitorCount(out uint pcRemoteMonitors);
        new HRESULT GetRemoteMonitorsBoundingBox(out int pLeft, out int pTop, out int pRight, out int pBottom);
        new HRESULT get_RemoteMonitorLayoutMatchesLocal(out VARIANT_BOOL pfRemoteMatchesLocal);
        new HRESULT put_DisableConnectionBar(VARIANT_BOOL _arg1);
        new HRESULT put_DisableRemoteAppCapsCheck(VARIANT_BOOL pfDisableRemoteAppCapsCheck);
        new HRESULT get_DisableRemoteAppCapsCheck(out VARIANT_BOOL pfDisableRemoteAppCapsCheck);
        new HRESULT put_WarnAboutDirectXRedirection(VARIANT_BOOL pfWarn);
        new HRESULT get_WarnAboutDirectXRedirection(out VARIANT_BOOL pfWarn);
        new HRESULT put_AllowPromptingForCredentials(VARIANT_BOOL pfAllow);
        new HRESULT get_AllowPromptingForCredentials(out VARIANT_BOOL pfAllow);

        new HRESULT SendLocation2D(double latitude, double longitude);
        new HRESULT SendLocation3D(double latitude, double longitude, int altitude);

        HRESULT get_CameraRedirConfigCollection(out IMsRdpCameraRedirConfigCollection ppCameraCollection);
        HRESULT DisableDpiCursorScalingForProcess();
        HRESULT get_Clipboard(out IMsRdpClipboard ppClipboard);
    }

    [ComImport]
    [Guid("ae45252b-aaab-4504-b681-649d6073a37a")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpCameraRedirConfigCollection
    {
        HRESULT Rescan();
        HRESULT get_Count(out uint pCount);
        HRESULT get_ByIndex(uint index, out IMsRdpCameraRedirConfig ppConfig);
        HRESULT get_BySymbolicLink(string SymbolicLink, out IMsRdpCameraRedirConfig ppConfig);
        HRESULT get_ByInstanceId(string InstanceId, out IMsRdpCameraRedirConfig ppConfig);
        HRESULT AddConfig(string SymbolicLink, VARIANT_BOOL fRedirected);
        HRESULT put_RedirectByDefault(VARIANT_BOOL pfRedirect);
        HRESULT get_RedirectByDefault(out VARIANT_BOOL pfRedirect);
        HRESULT put_EncodeVideo(VARIANT_BOOL pfEncode);
        HRESULT get_EncodeVideo(out VARIANT_BOOL pfEncode);
        HRESULT put_EncodingQuality(CameraRedirEncodingQuality pEncodingQuality);
        HRESULT get_EncodingQuality(out CameraRedirEncodingQuality pEncodingQuality);
    }

    [ComImport]
    [Guid("09750604-d625-47c1-9fcd-f09f735705d7")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpCameraRedirConfig
    {
        HRESULT get_FriendlyName(out string pFriendlyName);
        HRESULT get_SymbolicLink(out string pSymbolicLink);
        HRESULT get_InstanceId(out string pInstanceId);
        HRESULT get_ParentInstanceId(out string pParentInstanceId);
        HRESULT put_Redirected(VARIANT_BOOL pfRedirected);
        HRESULT get_Redirected(out VARIANT_BOOL pfRedirected);
        HRESULT get_DeviceExists(out VARIANT_BOOL pfExists);
    }

    public enum CameraRedirEncodingQuality
    {
        encodingQualityLow = 0,
        encodingQualityMedium = 1,
        encodingQualityHigh = 2
    };

    [ComImport]
    [Guid("2e769ee8-00c7-43dc-afd9-235d75b72a40")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMsRdpClipboard
    {
        HRESULT CanSyncLocalClipboardToRemoteSession(out VARIANT_BOOL pfSync);
        HRESULT SyncLocalClipboardToRemoteSession();
        HRESULT CanSyncRemoteClipboardToLocalSession(out VARIANT_BOOL pfSync);
        HRESULT SyncRemoteClipboardToLocalSession();
    }

    // Copied from Interop DLL
    [Guid("336D5562-EFA8-482E-8CB3-C5C0FC7A7DB6"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType(TypeLibTypeFlags.FDispatchable)]
    [ComImport]
    public interface IMsTscAxEvents
    {
        [DispId(1)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnConnecting();

        [DispId(2)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnConnected();

        [DispId(3)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnLoginComplete();

        [DispId(4)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnDisconnected([In] int discReason);

        [DispId(5)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnEnterFullScreenMode();

        [DispId(6)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnLeaveFullScreenMode();

        [DispId(7)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnChannelReceivedData([MarshalAs(UnmanagedType.BStr)][In] string chanName, [MarshalAs(UnmanagedType.BStr)][In] string data);

        [DispId(8)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnRequestGoFullScreen();

        [DispId(9)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnRequestLeaveFullScreen();

        [DispId(10)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnFatalError([In] int errorCode);

        [DispId(11)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnWarning([In] int warningCode);

        [DispId(12)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnRemoteDesktopSizeChange([In] int width, [In] int height);

        [DispId(13)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnIdleTimeoutNotification();

        [DispId(14)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnRequestContainerMinimize();

        [DispId(15)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        bool OnConfirmClose();

        [DispId(16)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        bool OnReceivedTSPublicKey([MarshalAs(UnmanagedType.BStr)][In] string publicKey);

        [DispId(17)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        [return: ComAliasName("MSTSCLib.AutoReconnectContinueState")]
        AutoReconnectContinueState OnAutoReconnecting([In] int disconnectReason, [In] int attemptCount);

        [DispId(18)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnAuthenticationWarningDisplayed();

        [DispId(19)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnAuthenticationWarningDismissed();

        [DispId(20)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnRemoteProgramResult([MarshalAs(UnmanagedType.BStr)][In] string bstrRemoteProgram, [ComAliasName("MSTSCLib.RemoteProgramResult")][In] RemoteProgramResult lError, [In] bool vbIsExecutable);

        [DispId(21)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnRemoteProgramDisplayed([In] bool vbDisplayed, [In] uint uDisplayInformation);

        [DispId(29)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnRemoteWindowDisplayed([In] bool vbDisplayed, [ComAliasName("MSTSCLib.wireHWND")][In] ref IntPtr hwnd, [ComAliasName("MSTSCLib.RemoteWindowDisplayedAttribute")][In] RemoteWindowDisplayedAttribute windowAttribute);

        [DispId(22)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnLogonError([In] int lError);

        [DispId(23)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnFocusReleased([In] int iDirection);

        [DispId(24)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnUserNameAcquired([MarshalAs(UnmanagedType.BStr)][In] string bstrUserName);

        [DispId(26)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnMouseInputModeChanged([In] bool fMouseModeRelative);

        [DispId(28)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnServiceMessageReceived([MarshalAs(UnmanagedType.BStr)][In] string serviceMessage);

        [DispId(30)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnConnectionBarPullDown();

        [DispId(32)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnNetworkStatusChanged([In] uint qualityLevel, [In] int bandwidth, [In] int rtt);

        [DispId(35)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnDevicesButtonPressed();

        [DispId(33)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnAutoReconnected();

        [DispId(34)]
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        void OnAutoReconnecting2([In] int disconnectReason, [In] bool networkAvailable, [In] int attemptCount, [In] int maxAttemptCount);
    }

    public enum AutoReconnectContinueState
    {
        autoReconnectContinueAutomatic = 0,
        autoReconnectContinueStop = 1,
        autoReconnectContinueManual = 2
    };

    public enum RemoteProgramResult
    {
        remoteAppResultOk = 0,
        remoteAppResultLocked = 1,
        remoteAppResultProtocolError = 2,
        remoteAppResultNotInWhitelist = 3,
        remoteAppResultNetworkPathDenied = 4,
        remoteAppResultFileNotFound = 5,
        remoteAppResultFailure = 6,
        remoteAppResultHookNotLoaded = 7
    };

    public enum RemoteWindowDisplayedAttribute
    {
        remoteAppWindowNone = 0,
        remoteAppWindowDisplayed = 1,
        remoteAppShellIconDisplayed = 2
    };
    
    public class CEventSink : IMsTscAxEvents
    {
        MainWindow _mw = null;

        public CEventSink(MainWindow mw)
        {
            _mw = mw;
        }

        public int GetTypeInfoCount()
        {
            return 0;
        }

        [return: MarshalAs(UnmanagedType.Interface)]
        public IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid)
        {
            return IntPtr.Zero;
        }

        public HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames, [In, MarshalAs(UnmanagedType.U4)] int lcid, [MarshalAs(UnmanagedType.LPArray), Out] int[] rgDispId)
        {
            return HRESULT.S_OK;
        }

        public HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags, [In, Out] DISPPARAMS pDispParams, [Out] out object pVarResult, [In, Out] EXCEPINFO pExcepInfo, [MarshalAs(UnmanagedType.LPArray), Out] IntPtr[] pArgErr)
        {
            HRESULT hr = HRESULT.S_OK;
            switch (dispIdMember)
            {
                case MSRDPTools.DISPID_CONNECTED:
                    OnConnected();
                    break;
            }
            pVarResult = 1;
            return hr;
        }


        public void OnConnecting()
        {
            _mw.OnConnecting();
            return;
        }

        public void OnConnected()
        {
            _mw.OnConnected();
            return;
        }

        public void OnLoginComplete()
        {
            return;
        }

        public void OnDisconnected([In] int discReason)
        {
            _mw.OnDisconnected();
            return;
        }

        public void OnEnterFullScreenMode()
        {
            _mw.OnEnterFullScreenMode();
            return;
        }

        public void OnLeaveFullScreenMode()
        {
            _mw.OnLeaveFullScreenMode();
            return;
        }

        public void OnChannelReceivedData([In, MarshalAs(UnmanagedType.BStr)] string chanName, [In, MarshalAs(UnmanagedType.BStr)] string data)
        {
            return;
        }

        public void OnRequestGoFullScreen()
        {
            return;
        }

        public void OnRequestLeaveFullScreen()
        {
            return;
        }

        public void OnFatalError([In] int errorCode)
        {
            return;
        }

        public void OnWarning([In] int warningCode)
        {
            return;
        }

        public void OnRemoteDesktopSizeChange([In] int width, [In] int height)
        {
            return;
        }

        public void OnIdleTimeoutNotification()
        {
            return;
        }

        public void OnRequestContainerMinimize()
        {
            return;
        }

        public bool OnConfirmClose()
        {
            return true;
        }

        public bool OnReceivedTSPublicKey([In, MarshalAs(UnmanagedType.BStr)] string publicKey)
        {
            return true;
        }

        public AutoReconnectContinueState OnAutoReconnecting([In] int disconnectReason, [In] int attemptCount)
        {
            return AutoReconnectContinueState.autoReconnectContinueAutomatic;
        }

        public void OnAuthenticationWarningDisplayed()
        {
            return;
        }

        public void OnAuthenticationWarningDismissed()
        {
            return;
        }

        public void OnRemoteProgramResult([In, MarshalAs(UnmanagedType.BStr)] string bstrRemoteProgram, [In] RemoteProgramResult lError, [In] bool vbIsExecutable)
        {
            return;
        }

        public void OnRemoteProgramDisplayed([In] bool vbDisplayed, [In] uint uDisplayInformation)
        {
            return;
        }

        public void OnRemoteWindowDisplayed([In] bool vbDisplayed, [In] ref IntPtr hwnd, [In] RemoteWindowDisplayedAttribute windowAttribute)
        {
            return;
        }

        public void OnLogonError([In] int lError)
        {
            return;
        }

        public void OnFocusReleased([In] int iDirection)
        {
            return;
        }

        public void OnUserNameAcquired([In, MarshalAs(UnmanagedType.BStr)] string bstrUserName)
        {
            return;
        }

        public void OnMouseInputModeChanged([In] bool fMouseModeRelative)
        {
            return;
        }

        public void OnServiceMessageReceived([In, MarshalAs(UnmanagedType.BStr)] string serviceMessage)
        {
            return;
        }

        public void OnConnectionBarPullDown()
        {
            return;
        }

        public void OnNetworkStatusChanged([In] uint qualityLevel, [In] int bandwidth, [In] int rtt)
        {
            return;
        }

        public void OnDevicesButtonPressed()
        {
            return;
        }

        public void OnAutoReconnected()
        {
            return;
        }

        public void OnAutoReconnecting2([In] int disconnectReason, [In] bool networkAvailable, [In] int attemptCount, [In] int maxAttemptCount)
        {
            return;
        }
    }

    public class MSRDPTools
    {
        public static Guid CLSID_MsRdpClient11NotSafeForScripting = new Guid(0x1df7c823, 0xb2d4, 0x4b54, 0x97, 0x5a, 0xf2, 0xac, 0x5d, 0x7c, 0xf8, 0xb8);
        public static Guid CLSID_MsRdpClient10NotSafeForScripting = new Guid(0xa0c63c30, 0xf08d, 0x4ab4, 0x90, 0x7c, 0x34, 0x90, 0x5d, 0x77, 0x0c, 0x7d);
        public static Guid CLSID_MsRdpClient9NotSafeForScripting = new Guid(0x8b918b82, 0x7985, 0x4c24, 0x89, 0xdf, 0xc3, 0x3a, 0xd2, 0xbb, 0xfb, 0xcd);
        public static Guid CLSID_MsRdpClient8NotSafeForScripting = new Guid(0xa3bc03a0, 0x041d, 0x42e3, 0xad, 0x22, 0x88, 0x2b, 0x78, 0x65, 0xc9, 0xc5);
        public static Guid CLSID_MsRdpClient7NotSafeForScripting = new Guid(0x54d38bf7, 0xb1ef, 0x4479, 0x96, 0x74, 0x1b, 0xd6, 0xea, 0x46, 0x52, 0x58);
        public static Guid CLSID_MsRdpClient6NotSafeForScripting = new Guid(0xd2ea46a7, 0xc2bf, 0x426b, 0xaf, 0x24, 0xe1, 0x9c, 0x44, 0x45, 0x63, 0x99);
        public static Guid CLSID_MsRdpClient5NotSafeForScripting = new Guid(0x4eb2f086, 0xc818, 0x447e, 0xb3, 0x2c, 0xc5, 0x1c, 0xe2, 0xb3, 0x0d, 0x31);
        public static Guid CLSID_MsRdpClient4NotSafeForScripting = new Guid(0x6ae29350, 0x321b, 0x42be, 0xbb, 0xe5, 0x12, 0xfb, 0x52, 0x70, 0xc0, 0xde);
        public static Guid CLSID_MsRdpClient3NotSafeForScripting = new Guid(0xace575fd, 0x1fcf, 0x4074, 0x94, 0x01, 0xeb, 0xab, 0x99, 0x0f, 0xa9, 0xde);
        public static Guid CLSID_MsRdpClient2NotSafeForScripting = new Guid(0x3523c2fb, 0x4031, 0x44e4, 0x9a, 0x3b, 0xf1, 0xe9, 0x49, 0x86, 0xee, 0x7f);
        public static Guid CLSID_MsRdpClientNotSafeForScripting = new Guid(0x7cacbd7b, 0x0d99, 0x468f, 0xac, 0x33, 0x22, 0xe4, 0x95, 0xc0, 0xaf, 0xe5);

        public const int DISPID_CONNECTING = (1);
        public const int DISPID_CONNECTED = (2);
        public const int DISPID_LOGINCOMPLETE = (3);
        public const int DISPID_DISCONNECTED = (4);
        public const int DISPID_ENTERFULLSCREENMODE = (5);
        public const int DISPID_LEAVEFULLSCREENMODE = (6);
        public const int DISPID_CHANNELRECEIVEDDATA = (7);
        public const int DISPID_REQUESTGOFULLSCREEN = (8);
        public const int DISPID_REQUESTLEAVEFULLSCREEN = (9);
        public const int DISPID_FATALERROR = (10);
        public const int DISPID_WARNING = (11);
        public const int DISPID_REMOTEDESKTOPSIZECHANGE = (12);
        public const int DISPID_IDLETIMEOUTNOTIFICATION = (13);
        public const int DISPID_REQUESTCONTAINERMINIMIZE = (14);
        public const int DISPID_CONFIRMCLOSE = (15);
        public const int DISPID_RECEVIEDTSPUBLICKEY = (16);
        public const int DISPID_AUTORECONNECTING = (17);
        public const int DISPID_INTERNALDIALOGDISPLAYED = (18);
        public const int DISPID_INTERNALDIALOGDISMISSED = (19);
        public const int DISPID_ONREMOTEPROGRAMRESULT = (20);
        public const int DISPID_ONREMOTEPROGRAMDISPLAYED = (21);
        public const int DISPID_LOGONERROR = (22);
        public const int DISPID_FOCUSRELEASED = (23);
        public const int DISPID_USERNAMEACQUIRED = (24);
        public const int DISPID_CHANNELWRITECOMPLETED = (25);
        public const int DISPID_MOUSEINPUTMODECHANGED = (26);
        public const int DISPID_ONSTATUSINFO = (27);
        public const int DISPID_SERVICEMESSAGERECEIVED = (28);
        public const int DISPID_ONREMOTEWINDOWDISPLAYED = (29);
        public const int DISPID_ONCONNECTIONBARPULLDOWN = (30);
        public const int DISPID_ONNETWORKSTATUSCHANGED = (32);
        public const int DISPID_ONAUTORECONNECTED = (33);
        public const int DISPID_ONAUTORECONNECTING2 = (34);
        public const int DISPID_ONDEVICESBUTTONPRESSED = (35);

    }


}
