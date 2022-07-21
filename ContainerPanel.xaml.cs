using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace WinUI3_ActiveX_MSRDP
{
    public sealed partial class ContainerPanel : UserControl
    {
        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateWindowEx(int dwExStyle, string lpClassName, string lpWindowName, int dwStyle, int x, int y, int nWidth, int nHeight, IntPtr hWndParent, IntPtr hMenu, IntPtr hInstance, IntPtr lpParam);

        public const int WS_OVERLAPPED = 0x0;
        public const int WS_BORDER = 0x00800000;
        public const int WS_POPUP = unchecked((int)0x80000000L);
        public const int WS_CHILD = 0x40000000;
        public const int WS_MINIMIZE = 0x20000000;
        public const int WS_VISIBLE = 0x10000000;
        public const int WS_DISABLED = 0x8000000;

        public const int WS_EX_DLGMODALFRAME = 0x00000001;
        public const int WS_EX_NOPARENTNOTIFY = 0x00000004;
        public const int WS_EX_TOPMOST = 0x00000008;
        public const int WS_EX_ACCEPTFILES = 0x00000010;
        public const int WS_EX_TRANSPARENT = 0x00000020;
        public const int WS_EX_MDICHILD = 0x00000040;
        public const int WS_EX_TOOLWINDOW = 0x00000080;
        public const int WS_EX_WINDOWEDGE = 0x00000100;
        public const int WS_EX_CLIENTEDGE = 0x00000200;
        public const int WS_EX_CONTEXTHELP = 0x00000400;
        public const int WS_EX_RIGHT = 0x00001000;
        public const int WS_EX_LEFT = 0x00000000;
        public const int WS_EX_RTLREADING = 0x00002000;
        public const int WS_EX_LTRREADING = 0x00000000;
        public const int WS_EX_LEFTSCROLLBAR = 0x00004000;
        public const int WS_EX_RIGHTSCROLLBAR = 0x00000000;
        public const int WS_EX_CONTROLPARENT = 0x00010000;
        public const int WS_EX_STATICEDGE = 0x00020000;
        public const int WS_EX_APPWINDOW = 0x00040000;
        public const int WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE);
        public const int WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE | WS_EX_TOOLWINDOW | WS_EX_TOPMOST);
        public const int WS_EX_LAYERED = 0x00080000;
        public const int WS_EX_NOINHERITLAYOUT = 0x00100000; // Disable inheritence of mirroring by children
        public const int WS_EX_NOREDIRECTIONBITMAP = 0x00200000;
        public const int WS_EX_LAYOUTRTL = 0x00400000; // Right to left mirroring
        public const int WS_EX_COMPOSITED = 0x02000000;
        public const int WS_EX_NOACTIVATE = 0x08000000;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool MoveWindow(IntPtr hWnd, int x, int y, int cx, int cy, bool repaint);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateCompatibleDC(IntPtr hDC);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool DeleteObject(IntPtr hObject);

        public const int SRCCOPY = 0x00CC0020; /* dest = source */

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool BitBlt(IntPtr hDC, int x, int y, int nWidth, int nHeight, IntPtr hSrcDC, int xSrc, int ySrc, int dwRop);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateDIBSection(IntPtr hdc, ref BITMAPINFO pbmi, uint usage, ref IntPtr ppvBits, IntPtr hSection, int offset);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateDIBSection(IntPtr hdc, ref BITMAPV5HEADER pbmi, uint usage, ref IntPtr ppvBits, IntPtr hSection, int offset);

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFOHEADER
        {
            [MarshalAs(UnmanagedType.I4)]
            public int biSize;
            [MarshalAs(UnmanagedType.I4)]
            public int biWidth;
            [MarshalAs(UnmanagedType.I4)]
            public int biHeight;
            [MarshalAs(UnmanagedType.I2)]
            public short biPlanes;
            [MarshalAs(UnmanagedType.I2)]
            public short biBitCount;
            [MarshalAs(UnmanagedType.I4)]
            public int biCompression;
            [MarshalAs(UnmanagedType.I4)]
            public int biSizeImage;
            [MarshalAs(UnmanagedType.I4)]
            public int biXPelsPerMeter;
            [MarshalAs(UnmanagedType.I4)]
            public int biYPelsPerMeter;
            [MarshalAs(UnmanagedType.I4)]
            public int biClrUsed;
            [MarshalAs(UnmanagedType.I4)]
            public int biClrImportant;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFO
        {
            [MarshalAs(UnmanagedType.Struct, SizeConst = 40)]
            public BITMAPINFOHEADER bmiHeader;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1024)]
            public int[] bmiColors;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPV5HEADER
        {
            public int bV5Size;
            public int bV5Width;
            public int bV5Height;
            public short bV5Planes;
            public short bV5BitCount;
            public int bV5Compression;
            public int bV5SizeImage;
            public int bV5XPelsPerMeter;
            public int bV5YPelsPerMeter;
            public int bV5ClrUsed;
            public int bV5ClrImportant;
            public int bV5RedMask;
            public int bV5GreenMask;
            public int bV5BlueMask;
            public int bV5AlphaMask;
            public int bV5CSType;
            public CIEXYZTRIPLE bV5Endpoints;
            public int bV5GammaRed;
            public int bV5GammaGreen;
            public int bV5GammaBlue;
            public int bV5Intent;
            public int bV5ProfileData;
            public int bV5ProfileSize;
            public int bV5Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CIEXYZTRIPLE
        {
            public CIEXYZ ciexyzRed;
            public CIEXYZ ciexyzGreen;
            public CIEXYZ ciexyzBlue;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CIEXYZ
        {
            public int ciexyzX;
            public int ciexyzY;
            public int ciexyzZ;
        }

        public const int BI_RGB = 0;
        public const int BI_RLE8 = 1;
        public const int BI_RLE4 = 2;
        public const int BI_BITFIELDS = 3;
        public const int BI_JPEG = 4;
        public const int BI_PNG = 5;

        public const int DIB_RGB_COLORS = 0;
        public const int DIB_PAL_COLORS = 1;
        
        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

        public const int PW_CLIENTONLY = 0x1;
        public const int PW_RENDERFULLCONTENT = 0x2;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool ShowWindow(IntPtr hWnd, int nShowCmd);

        public const int SW_HIDE = 0;
        public const int SW_SHOWNORMAL = 1;
        public const int SW_SHOWMINIMIZED = 2;
        public const int SW_SHOWMAXIMIZED = 3;
        public const int SW_SHOWNOACTIVATE = 4;
        public const int SW_SHOW = 5;

        public const uint LWA_COLORKEY = 0x00000001;
        public const uint LWA_ALPHA = 0x00000002;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);



        private IntPtr _hWndParent = IntPtr.Zero;
        public IntPtr hWndContainer = IntPtr.Zero;

        public ContainerPanel()
        {
            this.InitializeComponent();
            this.SizeChanged += ContainerPanel_SizeChanged;
        }

        private void ContainerPanel_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (hWndContainer != IntPtr.Zero)
            {
                uint nDPI = GetDpiForWindow(_hWndParent);
                int nWidth = (int)(e.NewSize.Width * nDPI/96.0f);
                int nHeight = (int)(e.NewSize.Height * nDPI/96.0f);
                MoveWindow(hWndContainer, (int)this.ActualOffset.X, (int)this.ActualOffset.Y, nWidth, nHeight, true);
            }
        }
   
        public void ContainerPanelInit(string sClass, string sText, IntPtr hWndParent)
        {
            //var currentMainWindow = ((App)Application.Current).m_window;
            if (hWndParent != IntPtr.Zero)
            {
                _hWndParent = hWndParent;
                hWndContainer = CreateWindowEx(WS_EX_LAYERED, sClass, sText, WS_VISIBLE | WS_CHILD, 0, 0, (int)this.ActualWidth, (int)this.ActualHeight, _hWndParent, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                //hWndContainer = CreateWindowEx(WS_EX_TRANSPARENT, "Static", "", WS_VISIBLE | WS_CHILD | WS_BORDER, 0, 0, (int)this.ActualWidth, (int)this.ActualHeight, _hWndParent, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            }
        }

        public ImageSource Source
        {
            get { return (ImageSource)GetValue(SourceProperty); }
            set { SetValue(SourceProperty, value); }
        }    

        public static readonly DependencyProperty SourceProperty = 
            DependencyProperty.Register("Source", typeof(ImageSource),  typeof(ContainerPanel), null);

        public Stretch Stretch
        {
            get { return (Stretch)GetValue(StretchProperty); }
            set { SetValue(StretchProperty, value); }
        }

        public static readonly DependencyProperty StretchProperty =
            DependencyProperty.Register("Stretch", typeof(Stretch), typeof(ContainerPanel), null);

        public void Hide(bool bCopy)
        {
            if (bCopy)
                CopyImage();
            ShowWindow(hWndContainer, SW_HIDE);
        }

        public void Show(bool bReset)
        {
            if (bReset)
                Source = null;
            ShowWindow(hWndContainer, SW_SHOWNORMAL);
        }

        private async void CopyImage()
        { 
            IntPtr hDC = GetDC(IntPtr.Zero);
            IntPtr hDCMem = CreateCompatibleDC(hDC);
            int nX = (int)this.ActualWidth;
            int nY = (int)this.ActualHeight;
            BITMAPINFO bi = new BITMAPINFO();
            bi.bmiHeader.biSize = Marshal.SizeOf(typeof(BITMAPINFOHEADER));
            bi.bmiHeader.biWidth = nX;
            bi.bmiHeader.biHeight = -nY;
            bi.bmiHeader.biPlanes = 1;
            bi.bmiHeader.biBitCount = 32;
            bi.bmiHeader.biCompression = BI_RGB;
            IntPtr pBits = IntPtr.Zero;
            IntPtr hBitmap = CreateDIBSection(hDC, ref bi, DIB_RGB_COLORS, ref pBits, IntPtr.Zero, 0);
            if (hBitmap != IntPtr.Zero)
            {
                IntPtr hBitmapOld = SelectObject(hDCMem, hBitmap);
                PrintWindow(hWndContainer, hDCMem, PW_RENDERFULLCONTENT);
                //BitBlt(hDC, 0, 0, nX, nY, hDCMem, 0, 0, SRCCOPY);
                int nSize = nX * nY * 4;
                byte[] pManagedArray = new byte[nSize];
                Marshal.Copy(new IntPtr(pBits.ToInt32()), pManagedArray, 0, nSize);
                Microsoft.UI.Xaml.Media.Imaging.WriteableBitmap wb = new Microsoft.UI.Xaml.Media.Imaging.WriteableBitmap(nX, nY);
                await wb.PixelBuffer.AsStream().WriteAsync(pManagedArray, 0, pManagedArray.Length);
                this.Source = wb;

                SelectObject(hDCMem, hBitmapOld);
                DeleteObject(hBitmap);               
            }
            DeleteObject(hDCMem);
            ReleaseDC(IntPtr.Zero, hDC);
        }

        public void SetOpacity(IntPtr hWnd, int nOpacity)
        {
            SetLayeredWindowAttributes(hWnd, 0, (byte)(255 * nOpacity / 100), LWA_ALPHA);
        }
    }
}
