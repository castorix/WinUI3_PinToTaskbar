using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Windows.Foundation;
using Windows.Foundation.Collections;

using GlobalStructures;
using static GlobalStructures.GlobalTools;
using PinnedList;
using static PinnedList.PinnedListTools;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.CompilerServices;
using ABI.Windows.Foundation;
using System.Security.Claims;
using System.Text;
using WinRT.Interop;
using System.Runtime;
using Microsoft.UI.Xaml.Media.Imaging;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using Microsoft.UI;
using System.Diagnostics;
using Windows.Storage;
using System.ComponentModel;
using Windows.ApplicationModel.DataTransfer;
using Shell32;
using System.Xml;
using static System.Net.WebRequestMethods;
using System.Reflection;
using Microsoft.Win32.SafeHandles;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Reflection.Metadata;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace WinUI3_PinToTaskbar
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO piconinfo);

        [StructLayout(LayoutKind.Sequential)]
        public struct ICONINFO
        {
            public bool fIcon;
            public uint xHotspot;
            public uint yHotspot;
            public IntPtr hbmMask;
            public IntPtr hbmColor;
        }

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetObject(IntPtr hFont, int nSize, out BITMAP bm);

        [StructLayoutAttribute(LayoutKind.Sequential)]
        public struct BITMAP
        {
            public int bmType;
            public int bmWidth;
            public int bmHeight;
            public int bmWidthBytes;
            public short bmPlanes;
            public short bmBitsPixel;
            public IntPtr bmBits;
        }

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateCompatibleDC(IntPtr hDC);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool DeleteDC(IntPtr hDC);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth, int nHeight);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetDIBits(IntPtr hdc, IntPtr hbm, uint start, uint cLines, byte[] lpvBits, ref BITMAPINFO lpbmi, uint usage);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetDIBits(IntPtr hdc, IntPtr hbm, uint start, uint cLines, byte[] lpvBits, ref BITMAPV5HEADER lpbmi, uint usage);

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

        public delegate int SUBCLASSPROC(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, uint dwRefData);

        [DllImport("Comctl32.dll", SetLastError = true)]
        public static extern bool SetWindowSubclass(IntPtr hWnd, SUBCLASSPROC pfnSubclass, uint uIdSubclass, uint dwRefData);

        [DllImport("Comctl32.dll", SetLastError = true)]
        public static extern int DefSubclassProc(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam);

        [DllImport("User32.dll", SetLastError = true)]
        public static extern IntPtr GetShellWindow();

        [DllImport("User32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        
        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true)]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr LoadLibrary(string lpLibFileName);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateFileMapping(IntPtr hFile, IntPtr lpFileMappingAttributes, uint flProtect, uint dwMaximumSizeHigh, uint dwMaximumSizeLow, string lpName);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr MapViewOfFile(IntPtr hFileMappingObject, uint dwDesiredAccess, uint dwFileOffsetHigh, uint dwFileOffsetLow, int dwNumberOfBytesToMap);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool UnmapViewOfFile(IntPtr lpBaseAddress);

        [DllImport("Kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CloseHandle(IntPtr hObject);

        public const int INVALID_HANDLE_VALUE = -1;

        public const int SECTION_QUERY = 0x0001;
        public const int SECTION_MAP_WRITE = 0x0002;
        public const int SECTION_MAP_READ = 0x0004;
        public const int SECTION_MAP_EXECUTE = 0x0008;
        public const int SECTION_EXTEND_SIZE = 0x0010;
        public const int SECTION_MAP_EXECUTE_EXPLICIT = 0x0020; // not included in SECTION_ALL_ACCESS

        public const int STANDARD_RIGHTS_REQUIRED = 0x000F0000;
        public const int SECTION_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED | SECTION_QUERY | SECTION_MAP_WRITE | SECTION_MAP_READ |
            SECTION_MAP_EXECUTE | SECTION_EXTEND_SIZE);
        public const int FILE_MAP_WRITE = SECTION_MAP_WRITE;
        public const int FILE_MAP_READ = SECTION_MAP_READ;
        public const int FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS;

        public const int PAGE_NOACCESS = 0x01;
        public const int PAGE_READONLY = 0x02;
        public const int PAGE_READWRITE = 0x04;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool PostThreadMessage(uint idThread, uint Msg, IntPtr wParam, IntPtr lParam);


        IPinManagerInterop? m_pPMI = null;
        IStartMenuPinnedList? m_pSMPL = null;
        TaskbandPinClass? m_clsTB = null;

        private SUBCLASSPROC? SubClassDelegate;
        public const int WM_USER = 0x0400;
        private const int WM_SHELLNOTIFY  = WM_USER + 100;
        private const int WM_PINMESSAGE = WM_USER + 10;
        private const int WM_PINNOTIFY = WM_USER + 11;

        private IntPtr hWndMain = IntPtr.Zero;

        ObservableCollection<PinnedItem> pinnedItems = new ObservableCollection<PinnedItem>();
        public Windows.Media.Playback.MediaPlayer m_mp = new Windows.Media.Playback.MediaPlayer();
        private Windows.Media.Core.MediaSource m_UnpinSource;
        private Windows.Media.Core.MediaSource m_PinSource;

        public MainWindow()
        {
            this.InitializeComponent();
            hWndMain = WinRT.Interop.WindowNative.GetWindowHandle(this);
            this.Title = "WinUI 3 : Test Pin/Unpin to taskbar";
            App.Current.Resources["ButtonBackgroundPressed"] = App.Current.Resources["SystemAccentColor"];
            App.Current.Resources["ButtonBackgroundPointerOver"] = new SolidColorBrush(Colors.Blue);

            try
            {
                Type? PinManagerType = Type.GetTypeFromCLSID(CLSID_PinManager, true);
                m_pPMI = (IPinManagerInterop?)Activator.CreateInstance(PinManagerType);

                Type? StartMenuPinType = Type.GetTypeFromCLSID(CLSID_StartMenuPin, true);
                m_pSMPL = (IStartMenuPinnedList?)Activator.CreateInstance(StartMenuPinType);

                m_clsTB = new TaskbandPinClass();
            }
            catch (System.Exception ex)
            {
                Quit(ex);
                return;
            }           

            SubClassDelegate = new SUBCLASSPROC(WindowSubClass);
            bool bReturn = SetWindowSubclass(hWndMain, SubClassDelegate, 0, 0);
            SHChangeNotifyEntry cne = new SHChangeNotifyEntry();
            cne.fRecursive = false;
            cne.pidl = IntPtr.Zero;
            uint nNotify = SHChangeNotifyRegister(hWndMain, SHCNRF_NewDelivery | SHCNRF_ShellLevel, SHCNE_EXTENDED_EVENT, WM_SHELLNOTIFY, 1, ref cne);

            EnumPinnedItems();
            LoadMP3s();
            string? sExeName = Environment.ProcessPath;
            tbFile.Text = sExeName;            

            this.Closed += MainWindow_Closed;
        }

        private void MainWindow_Closed(object sender, WindowEventArgs args)
        {
            SafeRelease(ref m_pPMI);
            SafeRelease(ref m_pSMPL);
            //SafeRelease(ref m_clsTB);
            if (m_bHookInstalled)
            {
                InstallHook(false, 0);
            }
        }

        private async void Quit(System.Exception ex)
        {
            string sErrorMessage = ex.Message + Environment.NewLine + "Be sure to compile in x64";
            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
            WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
            _ = await md.ShowAsync();
            Application.Current.Exit();
        }

        private async void LoadMP3s()
        {
            await PreloadMP3s();
        }

        private async Task PreloadMP3s()
        {
            string basePath = AppContext.BaseDirectory;
            StorageFile unpinFile = await StorageFile.GetFileFromPathAsync(Path.Combine(basePath, "Assets\\Unpin.mp3"));
            StorageFile pinFile = await StorageFile.GetFileFromPathAsync(Path.Combine(basePath, "Assets\\Pin.mp3"));
            m_UnpinSource = Windows.Media.Core.MediaSource.CreateFromStorageFile(unpinFile);
            m_PinSource = Windows.Media.Core.MediaSource.CreateFromStorageFile(pinFile);
        }

        private void PlayUnpin()
        {
            m_mp.Source = m_UnpinSource;
            m_mp.Play();
        }

        private void PlayPin()
        {
            m_mp.Source = m_PinSource;
            m_mp.Play();
        }


        private async void EnumPinnedItems()
        {
            HRESULT hr = HRESULT.E_FAIL;
            IPinnedList3? pPL3 = null;
            IFlexibleTaskbarPinnedList? pFTPL = null;
            try
            {
                pPL3 = (IPinnedList3?)m_clsTB;              
            }
            catch (Exception ex)
            {
                pFTPL = (IFlexibleTaskbarPinnedList?)m_clsTB;
            }
            finally
            {
                IEnumFullIDList? pEFIL = null;
                if (pPL3 != null)
                    hr = pPL3.EnumObjects(out pEFIL);
                else if (pFTPL != null)
                    hr = pFTPL.EnumObjects(out pEFIL);
                if (hr == HRESULT.S_OK)
                {
                    // Sometimes items are in double ?!
                    //pinnedItems.Clear();
                    //Debug.WriteLine($"After Clear : {pinnedItems.Count}");

                    //DispatcherQueue.TryEnqueue(() =>
                    //{
                        lvPinnedItems.ItemsSource = null;
                        pinnedItems.Clear();
                        lvPinnedItems.ItemsSource = pinnedItems;
                    //});   

                    IntPtr pidl = IntPtr.Zero;
                    uint nCeltFetched = 0;
                    hr = pEFIL.Next(1, out pidl, out nCeltFetched);
                    while (hr == HRESULT.S_OK && (nCeltFetched) == 1)
                    {
                        string sDisplayName = "";
                        SHFILEINFO shfi = new SHFILEINFO();
                        IntPtr pRet = SHGetFileInfo(pidl, 0, ref shfi, (uint)Marshal.SizeOf(shfi), SHGFI_PIDL | SHGFI_DISPLAYNAME);
                        sDisplayName = shfi.szDisplayName;

                        WriteableBitmap? wb = null;
                        pRet = SHGetFileInfo(pidl, 0, ref shfi, (uint)Marshal.SizeOf(shfi), SHGFI_PIDL | SHGFI_ICON | SHGFI_LARGEICON | SHGFI_SYSICONINDEX);
                        if (pRet != IntPtr.Zero)
                        {
                            ICONINFO ii = new ICONINFO();
                            GetIconInfo(shfi.hIcon, out ii);
                            wb = await GetWriteableBitmapFromHBitmap(ii.hbmColor);
                        }

                        // C:\Users\Christian\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Control Panel.lnk
                        string? sPath = "";
                        IntPtr pszName;
                        // E_INVALIDARG
                        hr = SHGetNameFromIDList(pidl, SIGDN.SIGDN_FILESYSPATH, out pszName);
                        if (hr == HRESULT.S_OK)
                        {
                            sPath = Marshal.PtrToStringUni(pszName);
                            CoTaskMemFree(pszName);
                        }                       

                        string sLinkTarget = "";
                        IntPtr pidlChild = IntPtr.Zero;
                        IShellFolder? psf = null;                      
                        Guid IID_IShellFolder = typeof(IShellFolder).GUID;
                        hr = SHBindToParent(pidl, ref IID_IShellFolder, ref psf, ref pidlChild);
                        if (hr == HRESULT.S_OK)
                        {
                            uint rgfReserved = 0;
                            IShellLink? pShellLink = null;
                            Guid IID_IShellLink = typeof(IShellLink).GUID;
                            IntPtr pSHPtr = IntPtr.Zero;
                            hr = psf.GetUIObjectOf(IntPtr.Zero, 1, ref pidlChild, ref IID_IShellLink, ref rgfReserved, out pSHPtr);
                            if (hr == HRESULT.S_OK)
                            {
                                pShellLink = Marshal.GetObjectForIUnknown(pSHPtr) as IShellLink;
                                StringBuilder sbDisplayName = new StringBuilder(256);
                                WIN32_FIND_DATA wfd = new WIN32_FIND_DATA();
                                hr = pShellLink.GetPath(sbDisplayName, sbDisplayName.Capacity, ref wfd, SLGP_FLAGS.SLGP_RAWPATH);
                                sLinkTarget = sbDisplayName.ToString();

                                //StringBuilder sbIconLocation = new StringBuilder(260);
                                //int nIndex = 0;
                                //hr = pShellLink.GetIconLocation(sbIconLocation, sbIconLocation.Capacity, out nIndex);

                                SafeRelease(ref pShellLink);
                            }
                            else
                            {
                            }
                            SafeRelease(ref psf);                            
                        }

                        // "Microsoft.Windows.ControlPanel"
                        // shell:AppsFolder\Microsoft.Windows.ControlPanel
                        string sAppID = "";                        
                        IntPtr pAppId = Marshal.AllocHGlobal(256);
                        if (pPL3 != null)
                            hr = pPL3.GetAppIDForPinnedItem(pidl, out pAppId);
                        else if (pFTPL != null)
                            hr = pFTPL.GetAppIDForPinnedItem(pidl, out pAppId);
                        if (pAppId != IntPtr.Zero)
                            sAppID = Marshal.PtrToStringAuto(pAppId);
                        Marshal.FreeHGlobal(pAppId);

                        if (sLinkTarget == "")
                            sLinkTarget = "shell:AppsFolder\\" + sAppID;

                        if (!pinnedItems.Any(item => item.AppID == sAppID))
                        {
                            pinnedItems.Add(new PinnedItem(sDisplayName, sAppID, sLinkTarget, sPath, wb));
                        }

                        ILFree(pidl);
                        hr = pEFIL.Next(1, out pidl, out nCeltFetched);
                    }
                }
            }
        }
 
        private async void btnPinPinManager_Click(object sender, RoutedEventArgs e)
        {
            if (m_pPMI != null)
            {
                string? sExeName = tbFile.Text;
                string sSuffix = Path.GetExtension(sExeName);
                if (sSuffix == ".lnk" || sSuffix == ".exe")
                {
                    IntPtr pidl = ILCreateFromPath(sExeName);
                    if (pidl != IntPtr.Zero)
                    {
                        string sLinkTarget = sExeName;
                        if (sSuffix == ".lnk")
                        {
                            Shell shell = new Shell();
                            Folder folder = shell.NameSpace(System.IO.Path.GetDirectoryName(sExeName));
                            FolderItem item = folder.ParseName(System.IO.Path.GetFileName(sExeName));
                            if (item != null)
                            {
                                ShellLinkObject link = (ShellLinkObject)item.GetLink;
                                sLinkTarget = link.Path;
                            }
                        }
                        if (!pinnedItems.Any(item => item.LinkTarget == sLinkTarget))
                        //if (!pinnedItems.Any(item => item.Path == sExeName))
                        {
                            HRESULT hr = m_pPMI.PinItemToTaskbarShim(pidl, PINNEDLISTMODIFYCALLER.PMC_CONTEXTMENU);
                        }
                        else
                        {
                            string sErrorMessage = sExeName + " is already pinned.\r\nYou must unpin it first";
                            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
                            WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                            _ = await md.ShowAsync();
                        }
                    }
                    else
                    {
                        int nError = Marshal.GetLastWin32Error();
                        string sErrorMessage = "Could not get pidl for " + sExeName + "\r\n";
                        sErrorMessage += $"Error Code: 0x{nError:X8}\n";
                        string sExceptionMessage = new Win32Exception(nError).Message;
                        if (sExceptionMessage is not null && sExceptionMessage != "")
                        {
                            sErrorMessage += $"Description: {sExceptionMessage}";
                        }
                        else
                        {
                            sErrorMessage += "Description: Unknown error";
                        }
                        Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
                        WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                        _ = await md.ShowAsync();
                    }
                }
                else
                {
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog("File not pinnable", "Error");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                    _ = await md.ShowAsync();
                }
            }
        }

        private async void btnPinPinnedList_Click(object sender, RoutedEventArgs e)
        {
            HRESULT hr = HRESULT.E_FAIL;
            IPinnedList3? pPL3 = null;
            IFlexibleTaskbarPinnedList? pFTPL = null;
            try
            {
                pPL3 = (IPinnedList3?)m_clsTB;
            }
            catch (Exception ex)
            {
                pFTPL = (IFlexibleTaskbarPinnedList?)m_clsTB;
            }
            finally
            {
                string? sExeName = tbFile.Text;
                string sSuffix = Path.GetExtension(sExeName);
                if (sSuffix == ".lnk" || sSuffix == ".exe")
                {
                    IntPtr pidl = ILCreateFromPath(sExeName);
                    if (pidl != IntPtr.Zero)
                    {
                        string sLinkTarget = sExeName;
                        if (sSuffix == ".lnk")
                        {
                            Shell shell = new Shell();
                            Folder folder = shell.NameSpace(System.IO.Path.GetDirectoryName(sExeName));
                            FolderItem item = folder.ParseName(System.IO.Path.GetFileName(sExeName));
                            if (item != null)
                            {
                                ShellLinkObject link = (ShellLinkObject)item.GetLink;
                                sLinkTarget = link.Path;
                            }
                        }
                        if (!pinnedItems.Any(item => item.LinkTarget == sLinkTarget))
                        {
                            if (pPL3 != null)
                                hr = pPL3.Modify(IntPtr.Zero, pidl, PINNEDLISTMODIFYCALLER.PMC_CONTEXTMENU);
                            else if (pFTPL != null)
                                hr = pFTPL.Modify(IntPtr.Zero, pidl, PINNEDLISTMODIFYCALLER.PMC_CONTEXTMENU);
                            else
                            {
                                Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog("PinnedList interfaces not found", "Error");
                                WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                                _ = await md.ShowAsync();
                            }
                        }
                        else
                        {
                            string sErrorMessage = sExeName + " is already pinned.\r\nYou must unpin it first";
                            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
                            WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                            _ = await md.ShowAsync();
                        }
                    }
                    else
                    {
                        int nError = Marshal.GetLastWin32Error();
                        string sErrorMessage = "Could not get pidl for " + sExeName + "\r\n";
                        sErrorMessage += $"Error Code: 0x{nError:X8}\n";
                        string sExceptionMessage = new Win32Exception(nError).Message;
                        if (sExceptionMessage is not null && sExceptionMessage != "")
                        {
                            sErrorMessage += $"Description: {sExceptionMessage}";
                        }
                        else
                        {
                            sErrorMessage += "Description: Unknown error";
                        }
                        Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
                        WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                        _ = await md.ShowAsync();
                    }
                }
                else
                {
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog("File not pinnable", "Error");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                    _ = await md.ShowAsync();
                }
            }
        }

        private async void btnUnpin_Click(object sender, RoutedEventArgs e)
        {
            // https://learn.microsoft.com/en-us/windows/win32/api/shobjidl/nf-shobjidl-istartmenupinnedlist-removefromlist#examples
            // Windows 8: Unpins the item from the taskbar but does not remove the item from the Start screen. 

            if (lvPinnedItems.SelectedItem is PinnedItem selected)
            {
                string sAppID = selected.AppID;
                IShellItem? pShellItem = null;
                HRESULT hr = SHCreateItemFromParsingName(sAppID, IntPtr.Zero, typeof(IShellItem).GUID, out pShellItem);
                // 0x80070002  ERROR_FILE_NOT_FOUND
                if (hr == HRESULT.S_OK)
                {
                    if (m_pSMPL is not null)
                    {
                        hr = m_pSMPL.RemoveFromList(pShellItem);
                        if (hr == HRESULT.S_OK)
                            PlayUnpin();
                    }
                    SafeRelease(ref pShellItem);
                }
                else
                {
                    sAppID = selected.Path;                    
                    hr = SHCreateItemFromParsingName(sAppID, IntPtr.Zero, typeof(IShellItem).GUID, out pShellItem);
                    if (hr == HRESULT.S_OK)
                    {
                        if (m_pSMPL is not null)
                        {
                            hr = m_pSMPL.RemoveFromList(pShellItem);
                            if (hr == HRESULT.S_OK)
                                PlayUnpin();
                        }
                        SafeRelease(ref pShellItem);
                    }
                    else
                    {
                        sAppID = selected.AppID;
                        IntPtr pShellItemPtr = IntPtr.Zero;
                        hr = SHCreateItemInKnownFolder(FOLDERID_AppsFolder, 0, sAppID, typeof(IShellItem).GUID, out pShellItemPtr);
                        if (hr == HRESULT.S_OK)
                        {
                            pShellItem = Marshal.GetObjectForIUnknown(pShellItemPtr) as IShellItem;
                            if (pShellItem is not null)
                            {
                                if (m_pSMPL is not null)
                                {
                                    hr = m_pSMPL.RemoveFromList(pShellItem);
                                    if (hr == HRESULT.S_OK)
                                        PlayUnpin();
                                }
                                SafeRelease(ref pShellItem);
                            }
                        }
                        else
                        {
                            string sErrorMessage = $"Error Code: 0x{(int)hr:X8}\n";
                            Exception? ex = Marshal.GetExceptionForHR((int)hr);
                            if (ex != null)
                            {
                                sErrorMessage += $"Description: {ex.Message}";
                            }
                            else
                            {
                                sErrorMessage += "Description: Unknown error";
                            }
                            Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
                            WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                            _ = await md.ShowAsync();
                        }
                    }
                }
            }         
        }

        public static uint GetExplorerThreadId()
        {
            uint nExplorerPID = 0;
            uint nExplorerTID = 0;

            IntPtr hShellWindow = GetShellWindow();
            if (hShellWindow != IntPtr.Zero)
            {
                GetWindowThreadProcessId(hShellWindow, out nExplorerPID);
            }
            if (nExplorerPID == 0)
            {
                //Debug.WriteLine("Failed to get Explorer process ID.");
                return 0;
            }
            
            try
            {
                Process explorerProcess = Process.GetProcessById((int)nExplorerPID);
                ProcessThreadCollection threads = explorerProcess.Threads;
                if (threads.Count > 0)
                {
                    // Return the first thread ID found
                    nExplorerTID = (uint)threads[0].Id;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error accessing Explorer process: " + ex.Message);
            }
            return nExplorerTID;
        }

        private bool m_bHookInstalled = false;
        internal delegate bool SetHookDelegate(bool bInstall, uint dwThreadId);        

        bool InstallHook(bool bInstall, uint dwThreadId)
        {
            bool bRet = false;
            IntPtr hDll = LoadLibrary("Pin.dll");
            if (hDll != IntPtr.Zero)
            {
                IntPtr pFunc = GetProcAddress(hDll, "SetHook");
                if (pFunc != IntPtr.Zero)
                {                   
                    SetHookDelegate pSetHook = (SetHookDelegate)System.Runtime.InteropServices.Marshal.GetDelegateForFunctionPointer(pFunc, typeof(SetHookDelegate));
                    if (pSetHook != null)
                    {
                        try
                        {
                            bRet = pSetHook(bInstall, dwThreadId);
                        }
                        catch
                        {
                            // error
                        }                        
                    }
                }
                else
                {
                    // 127 ERROR_PROC_NOT_FOUND
                    int nError = Marshal.GetLastWin32Error();
                    bRet = false;
                }
            }
            return bRet;
        }

        private async void btnPinHook_Click(object sender, RoutedEventArgs e)
        {
            uint nExplorerThreadId = GetExplorerThreadId();
            if (!m_bHookInstalled)
            {               
                m_bHookInstalled = InstallHook(true, nExplorerThreadId);
                if (!m_bHookInstalled)
                {
                    int nError = Marshal.GetLastWin32Error();
                    string sErrorMessage = "Could not install Hook\r\n";
                    sErrorMessage += $"Error Code: 0x{nError:X8}\n";
                    string sExceptionMessage = new Win32Exception(nError).Message;
                    if (sExceptionMessage is not null && sExceptionMessage != "")
                    {
                        sErrorMessage += $"Description: {sExceptionMessage}";
                    }
                    else
                    {
                        sErrorMessage += "Description: Unknown error";
                    }
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                    _ = await md.ShowAsync();
                }
            }
            if (m_bHookInstalled)
            {
                IntPtr hMapFile = CreateFileMapping((IntPtr)INVALID_HANDLE_VALUE, IntPtr.Zero, PAGE_READWRITE, 0, 1024, "PinSharedMemory");
                if (hMapFile != IntPtr.Zero)
                {
                    IntPtr pSharedData = MapViewOfFile(hMapFile, FILE_MAP_ALL_ACCESS, 0, 0, 0);
                    if (pSharedData != IntPtr.Zero)
                    {
                        string? sExeName = tbFile.Text;
                        var sData = sExeName + '\0'; // Null-terminated like WCHAR*
                        byte[] bytes = System.Text.Encoding.Unicode.GetBytes(sData);
                        Marshal.Copy(bytes, 0, pSharedData, bytes.Length);
                        PostThreadMessage(nExplorerThreadId, WM_PINMESSAGE, 0, hWndMain);
                    }
                }
                else                
                {
                    int nError = Marshal.GetLastWin32Error();
                    string sErrorMessage = "CreateFileMapping failed !\r\n";
                    sErrorMessage += $"Error Code: 0x{nError:X8}\n";
                    string sExceptionMessage = new Win32Exception(nError).Message;
                    if (sExceptionMessage is not null && sExceptionMessage != "")
                    {
                        sErrorMessage += $"Description: {sExceptionMessage}";
                    }
                    else
                    {
                        sErrorMessage += "Description: Unknown error";
                    }
                    Windows.UI.Popups.MessageDialog md = new Windows.UI.Popups.MessageDialog(sErrorMessage, "Error");
                    WinRT.Interop.InitializeWithWindow.Initialize(md, hWndMain);
                    _ = await md.ShowAsync();
                }
            }
        }

        private void CopyLinkTarget_Click(object sender, RoutedEventArgs e)
        {           
            var menuItem = sender as MenuFlyoutItem;
            if (menuItem == null) return;
            
            var listViewItem = ((FrameworkElement)menuItem).DataContext as PinnedItem;
            if (listViewItem == null) return;
           
            var dataPackage = new DataPackage();
            dataPackage.SetText(listViewItem.LinkTarget);
            Clipboard.SetContent(dataPackage);
        }

        private async Task<string> OpenFileDialog()
        {
            var fop = new Windows.Storage.Pickers.FileOpenPicker();
            WinRT.Interop.InitializeWithWindow.Initialize(fop, hWndMain);
            //fop.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.PicturesLibrary;

            var types = new List<string> { ".exe", ".lnk" };
            foreach (var type in types)
                fop.FileTypeFilter.Add(type);

            // Crashes with shell:appsfolder or shell:::{35786D3C-B075-49b9-88DD-029876E11C01} or other CLSID
            try
            {
                var file = await fop.PickSingleFileAsync();
                return (file != null ? file.Path : string.Empty);
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }

        private async void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            string sFilePath = await OpenFileDialog();
            if (sFilePath != string.Empty)
            {
                tbFile.Text = sFilePath;
            }
        }

        public async Task<WriteableBitmap?> GetWriteableBitmapFromHBitmap(IntPtr hBitmap)
        {           
            if (hBitmap != IntPtr.Zero)
            {
                BITMAP bm;
                GetObject(hBitmap, Marshal.SizeOf(typeof(BITMAP)), out bm);
                int nWidth = bm.bmWidth;
                int nHeight = bm.bmHeight;
                BITMAPV5HEADER bi = new BITMAPV5HEADER();
                bi.bV5Size = Marshal.SizeOf(typeof(BITMAPV5HEADER));
                bi.bV5Width = nWidth;
                bi.bV5Height = -nHeight;
                bi.bV5Planes = 1;
                bi.bV5BitCount = 32;
                bi.bV5Compression = BI_BITFIELDS;
                bi.bV5AlphaMask = unchecked((int)0xFF000000);
                bi.bV5RedMask = 0x00FF0000;
                bi.bV5GreenMask = 0x0000FF00;
                bi.bV5BlueMask = 0x000000FF;

                IntPtr hDC = CreateCompatibleDC(IntPtr.Zero);
                IntPtr hBitmapOld = SelectObject(hDC, hBitmap);
                int nNumBytes = (int)(nWidth * 4 * nHeight);
                byte[] pPixels = new byte[nNumBytes];
                int nScanLines = GetDIBits(hDC, hBitmap, 0, (uint)nHeight, pPixels, ref bi, DIB_RGB_COLORS);                
                WriteableBitmap wb = new WriteableBitmap(nWidth, nHeight);
                await wb.PixelBuffer.AsStream().WriteAsync(pPixels, 0, pPixels.Length);
                SelectObject(hDC, hBitmapOld);
                DeleteDC(hDC);
                return wb;
            }
            else 
                return null;
        }

        private int WindowSubClass(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, uint dwRefData)
        {
            switch (uMsg)
            {
                case WM_SHELLNOTIFY:
                    { 
                        IntPtr hLock = SHChangeNotification_Lock(wParam, lParam, out IntPtr pppidl, out int plEvent);
                        if (hLock != IntPtr.Zero && plEvent == SHCNE_EXTENDED_EVENT)
                        {
                            IntPtr[] affectedPidls = new IntPtr[2];
                            Marshal.Copy(pppidl, affectedPidls, 0, affectedPidls.Length);                            
                            SHChangeDWORDAsIDList idl = new SHChangeDWORDAsIDList();
                            idl = (SHChangeDWORDAsIDList)Marshal.PtrToStructure(affectedPidls[0], typeof(SHChangeDWORDAsIDList));
                            // On Windows 10 : dwItem1 = 15 then 13 (twice when pin)
                            //if (idl.dwItem1 == SHCNEE_PINLISTCHANGED)
                            //{
                            //    Console.Beep(8000, 10);
                            //}
                            EnumPinnedItems();
                            SHChangeNotification_Unlock(hLock);
                        }
                    }
                    break;
                case WM_PINNOTIFY:
                    {
                        if (lParam == (IntPtr)HRESULT.S_OK)
                        {
                            if (wParam == (IntPtr)1)
                                PlayPin();
                            else if (wParam == (IntPtr)2)
                                PlayUnpin();
                        }
                    }
                    break;
            }
            return DefSubclassProc(hWnd, uMsg, wParam, lParam);
        }
    }

    public class SelectionToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            if (value is null)
                return false; 
            else
                return true;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }       
    }


    public class PinnedItem
    {
        #region Properties
        public string DisplayName { get; set; }
        public string AppID { get; set; }
        public string LinkTarget { get; set; }        
        public string? Path { get; set; }
        public WriteableBitmap Icon { get; set; }
        #endregion

        public PinnedItem(string sDisplayName, string sAppID, string sLinkTarget, string sPath, WriteableBitmap wbIcon)
        {
            DisplayName = sDisplayName;
            AppID = sAppID;
            LinkTarget = sLinkTarget;           
            Path = sPath;
            Icon = wbIcon;
        }
    }
}
