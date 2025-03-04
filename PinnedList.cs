using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.JavaScript;
using System.Security.Claims;
using System.Text;
using GlobalStructures;
using static WinUI3_PinToTaskbar.MainWindow;

namespace PinnedList
{
    internal class PinnedListTools
    {
        public static Guid CLSID_PinManager = new Guid("A5C8D635-B4ED-452B-8109-9501781096D1");
        public static Guid CLSID_TaskbandPin = new Guid("90AA3A4E-1CBA-4233-B8BB-535773D48449");
        public static Guid CLSID_StartMenuPin = new Guid("A2A9545D-A0C2-42B4-9708-A0B2BADD77C8");

        public static Guid FOLDERID_AppsFolder = new("1e87508d-89c2-42f0-8a7e-645a0f50ca58");

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr ILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern void ILFree(IntPtr pidl);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHCreateItemFromParsingName(string pszPath, IntPtr pbc, [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHCreateItemInKnownFolder([In, MarshalAs(UnmanagedType.LPStruct)] Guid kfid,
            uint dwKFFlags, string pszItem, [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHGetNameFromIDList(IntPtr pidl, SIGDN sigdnName, out IntPtr ppszName);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHCreateItemFromIDList(IntPtr pidl, [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);

        [DllImport("Ole32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern void CoTaskMemFree(IntPtr pv);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SHGetFileInfo(IntPtr pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        };

        public const int SHGFI_ICON = 0x000000100;     // get icon
        public const int SHGFI_DISPLAYNAME = 0x000000200;     // get display name
        public const int SHGFI_TYPENAME = 0x000000400;     // get type name
        public const int SHGFI_ATTRIBUTES = 0x000000800;     // get attributes
        public const int SHGFI_ICONLOCATION = 0x000001000;     // get icon location
        public const int SHGFI_EXETYPE = 0x000002000;     // return exe type
        public const int SHGFI_SYSICONINDEX = 0x000004000;     // get system icon index
        public const int SHGFI_LINKOVERLAY = 0x000008000;     // put a link overlay on icon
        public const int SHGFI_SELECTED = 0x000010000;     // show icon in selected state
        public const int SHGFI_ATTR_SPECIFIED = 0x000020000;     // get only specified attributes

        public const int SHGFI_LARGEICON = 0x000000000;     // get large icon
        public const int SHGFI_SMALLICON = 0x000000001;     // get small icon
        public const int SHGFI_OPENICON = 0x000000002;     // get open icon
        public const int SHGFI_SHELLICONSIZE = 0x000000004;     // get shell size icon
        public const int SHGFI_PIDL = 0x000000008;     // pszPath is a pidl
        public const int SHGFI_USEFILEATTRIBUTES = 0x000000010;     // use passed dwFileAttribute

        public const int SHGFI_ADDOVERLAYS = 0x000000020;     // apply the appropriate overlays
        public const int SHGFI_OVERLAYINDEX = 0x000000040;     // Get the index of the overlay in the upper 8 bits of the iIcon

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHBindToParent(IntPtr pidl, ref Guid riid, ref IShellFolder ppv, ref IntPtr ppidlLast);

        [DllImport("Shlwapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT StrRetToBuf(ref STRRET pstr, IntPtr pidl, StringBuilder pszBuf, [MarshalAs(UnmanagedType.U4)] uint cchBuf);

        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHGetDataFromIDList(IShellFolder psf, IntPtr pidl, int nFormat, out IntPtr pv, int cb);

        public const int SHGDFIL_FINDDATA = 1;
        public const int SHGDFIL_NETRESOURCE = 2;
        public const int SHGDFIL_DESCRIPTIONID = 3;

        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern uint SHChangeNotifyRegister(IntPtr hwnd, int fSources, int fEvents, uint wMsg, int cEntries, ref SHChangeNotifyEntry pshcne);

        public const int SHCNRF_InterruptLevel = 0x0001;
        public const int SHCNRF_ShellLevel = 0x0002;
        public const int SHCNRF_RecursiveInterrupt = 0x1000;
        public const int SHCNRF_ResumeThread = 0x2000;
        public const int SHCNRF_CreateSuspended = 0x4000;
        public const int SHCNRF_NewDelivery = 0x8000;

        public const int SHCNE_EXTENDED_EVENT = 0x04000000;
        public const int SHCNEE_PINLISTCHANGED = 10;

        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SHChangeNotification_Lock(IntPtr hChange, IntPtr dwProcId, out IntPtr pppidl, out int plEvent);

        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SHChangeNotification_Unlock(IntPtr hLock);  
    }

    [ComImport]
    [Guid("90AA3A4E-1CBA-4233-B8BB-535773D48449")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public class TaskbandPinClass
    {
    }

    [ComImport]
    [Guid("D75F625F-6DF9-4874-970D-17CBF846F00D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPinManagerInterop
    {
        HRESULT PinItemToTaskbarShim(IntPtr pidl, PINNEDLISTMODIFYCALLER p1);
        HRESULT PinItemFromTrustedCaller(IntPtr pidl, PINNEDLISTMODIFYCALLER p1);
        HRESULT ApplyPrependDefaultTaskbarLayout();
        HRESULT ApplyInPlaceTaskbarLayout(uint p0);
        HRESULT ApplyReorderTaskbarLayout(uint p0, uint p1);
    }

    // PinnedList interfaces adapted from https://github.com/explorer7-team/source/blob/a0add1196e1f1e31ebf91ace4674494631b16974/explorerwrapper/PinnedList.h#L151

    // IPinManagerInterop2 not on my Windows 10 22H2 OS
    //[ComImport]
    //[Guid("87d9e034-56d0-4f8c-be59-997b01754710")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //public interface IPinManagerInterop2 : IPinManagerInterop
    //{
    //    #region IPinManagerInterop
    //    new HRESULT PinItemToTaskbarShim(IntPtr pidl, PINNEDLISTMODIFYCALLER p1);
    //    new HRESULT PinItemFromTrustedCaller(IntPtr pidl, PINNEDLISTMODIFYCALLER p1);
    //    new HRESULT ApplyPrependDefaultTaskbarLayout();
    //    new HRESULT ApplyInPlaceTaskbarLayout(uint p0);
    //    new HRESULT ApplyReorderTaskbarLayout(uint p0, uint p1);
    //    #endregion

    //    HRESULT UnpinTaskbarItem(IntPtr pidl, PINNEDLISTMODIFYCALLER p1);
    //    HRESULT UpdatePinnedTaskbarItem(IntPtr pidl1, IntPtr pidl2, PINNEDLISTMODIFYCALLER p1);
    //}

    // from https://github.com/ntdiff/headers/blob/ebe89c140e89b475005cd8696597ddf7406dfb8a/Win11_2409_24H2/x64/System32/ole32.dll/Standalone/PINNEDLISTMODIFYCALLER.h#L4

    public enum PINNEDLISTMODIFYCALLER
    {
        PMC_APPRESOLVERMIGRATION = 0,
        PMC_APPRESOLVERUNINSTALL = 1,
        PMC_APPRESOLVERUNPINUNIQUEID = 2,
        PMC_CONTENTDELIVERYMANAGERBROKER = 3,
        PMC_CONTEXTMENU = 4,
        PMC_DEFAULTMFUCHANGE = 5,
        PMC_DEFAULTMFUPIN = 6,
        PMC_DEFAULTMFUPINAUX = 7,
        PMC_DEFAULTMFUTRYPIN = 8,
        PMC_DEFAULTMFUUPGRADE = 9,
        PMC_IEXPLORERCOMMAND = 10,
        PMC_JUMPVIEWBROKER = 11,
        PMC_PINNEDLISTLAYOUT = 12,
        PMC_PINNEDLISTNONEXIST = 13,
        PMC_PINNEDLISTREORDERLAYOUT = 14,
        PMC_PINNEDLISTUNRESOLVE = 15,
        PMC_RETAILDEMO = 16,
        PMC_SHELLLINK = 17,
        PMC_STARTMENU = 18,
        PMC_STARTMNU = 19,
        PMC_TASKBANDBADSHORTCUT = 20,
        PMC_TASKBANDBROKENPIN = 21,
        PMC_TASKBANDDEDUPPIN = 22,
        PMC_TASKBANDINSERT = 23,
        PMC_TASKBANDMODIFY = 24,
        PMC_TASKBANDPIN = 25,
        PMC_TASKBANDPINGROUP = 26,
        PMC_TASKBANDREORDER = 27,
        PMC_TASKBARPINNABLESURFACEBROKER = 28,
        PMC_TASKBARPINNABLESURFACEBROKERMIGRATION = 29,
        PMC_TASKBARPINNINGBROKERFACTORY = 30,
        PMC_TESTCODE = 31,
        PMC_UNIFIEDTILEMODELBROKER = 32,
        PMC_TRIMOOBEPINS = 33,
        PMC_MICROSOFTEDGE = 34,
        PMC_UPDATESHORTCUT = 35,
        PMC_BROWSERDECLUTTER = 36,
        PMC_CLOUDDEFAULTLAYOUT = 37,
        PMC_TASKBANDRESURRECTPLACEHOLDER = 38,
        PMC_TASKBANDTOMBSTONEPLACEHOLDER = 39,
    }

    [ComImport]
    [Guid("60274FA2-611F-4B8A-A293-F27BF103D148")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFlexibleTaskbarPinnedList
    {
        HRESULT EnumObjects(out IEnumFullIDList ppenumIDList);
        HRESULT GetPinnableInfo(IDataObject dataObject, uint dwPinnableFlag, out IShellItem2 shellItem2, out IShellItem shellItem,
             out IntPtr pAppID, out int pPinnable);
        HRESULT IsPinnable(IDataObject pdto, uint dwPinnableFlag, out IntPtr ppidl);
        HRESULT Resolve(IntPtr hwnd, uint dwFlags, IntPtr pidl, out IntPtr ppidlResolved);
        HRESULT Modify(IntPtr pidlFrom, IntPtr pidlTo, PINNEDLISTMODIFYCALLER plmCaller);
        HRESULT GetChangeCount(out uint uCount);
        HRESULT IsPinned(IntPtr pidl);
        HRESULT GetPinnedItem(IntPtr pidl, out IntPtr pinnedPidl);
        HRESULT GetAppIDForPinnedItem(IntPtr pidl, out IntPtr pAppID);
        HRESULT ItemChangeNotify(IntPtr pidlFrom, IntPtr pidlTo);
        HRESULT UpdateForRemovedItemsAsNecessary();
        HRESULT GetPinnedItemForAppID([MarshalAs(UnmanagedType.LPWStr)] string appId, out IntPtr pidl);
        HRESULT ApplyInPlaceTaskbarLayout(int param1, int param2);
        HRESULT ApplyReorderTaskbarLayout(int param1, int param2);
    }

    // Vista : IPinnedList : "C3C6EB6D-C837-4EAE-B172-5FEC52A2A4FD"
    // Windows 7, 8 : IPinnedList2 : "BBD20037-BC0E-42F1-913F-E2936BB0EA0C"
    [ComImport]
    [Guid("0DD79AE2-D156-45D4-9EEB-3B549769E940")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPinnedList3
    {
        HRESULT EnumObjects(out IEnumFullIDList ppenumIDList);
        HRESULT GetPinnableInfo(IDataObject dataObject, uint dwPinnableFlag, out IShellItem2 shellItem2, out IShellItem shellItem,
            out IntPtr pAppID, out int pPinnable);
        HRESULT IsPinnable(IDataObject pdto, uint dwPinnableFlag, out IntPtr ppidl);
        HRESULT Resolve(IntPtr hwnd, uint dwFlags, IntPtr pidl, out IntPtr ppidlResolved);
        HRESULT Unadvise(uint dwCookie);
        HRESULT GetChangeCount(out uint uCount);
        HRESULT IsPinned(IntPtr pidl);
        HRESULT GetPinnedItem(IntPtr pidl, out IntPtr pinnedPidl);
        HRESULT GetAppIDForPinnedItem(IntPtr pidl, out IntPtr pAppID);
        HRESULT ItemChangeNotify(IntPtr pidlFrom, IntPtr pidlTo);
        HRESULT UpdateForRemovedItemsAsNecessary();
        HRESULT PinShellLink(ushort us, IShellLink shellLinkW);
        HRESULT GetPinnedItemForAppID([MarshalAs(UnmanagedType.LPWStr)] string appId, out IntPtr pidl);
        HRESULT Modify(IntPtr pidlFrom, IntPtr pidlTo, PINNEDLISTMODIFYCALLER plmCaller);
    }

    [ComImport]
    [Guid("d0191542-7954-4908-bc06-b2360bbe45ba")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumFullIDList
    {
        [PreserveSig]
        HRESULT Next(uint celt, out IntPtr rgelt, out uint pceltFetched);
        [PreserveSig]
        HRESULT Skip(uint celt);
        [PreserveSig]
        HRESULT Reset();
        [PreserveSig]
        HRESULT Clone(out IEnumFullIDList ppenum);
    }

    [ComImport]
    [Guid("4CD19ADA-25A5-4A32-B3B7-347BEE5BE36B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStartMenuPinnedList
    {
        HRESULT RemoveFromList(IShellItem pitem);
    }

    [ComImport()]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
    public interface IShellItem
    {
        [PreserveSig()]
        HRESULT BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, ref IntPtr ppv);
        HRESULT GetParent(ref IShellItem ppsi);
        HRESULT GetDisplayName(SIGDN sigdnName, ref System.Text.StringBuilder ppszName);
        HRESULT GetAttributes(uint sfgaoMask, ref uint psfgaoAttribs);
        HRESULT Compare(IShellItem psi, uint hint, ref int piOrder);
    }

    [ComImport()]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("7e9fb0d3-919f-4307-ab2e-9b1860310c93")]
    public interface IShellItem2 : IShellItem
    {
        #region IShellItem
        [PreserveSig()]
        new HRESULT BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, ref IntPtr ppv);
        new HRESULT GetParent(ref IShellItem ppsi);
        new HRESULT GetDisplayName(SIGDN sigdnName, ref System.Text.StringBuilder ppszName);
        new HRESULT GetAttributes(uint sfgaoMask, ref uint psfgaoAttribs);
        new HRESULT Compare(IShellItem psi, uint hint, ref int piOrder);
        #endregion

        HRESULT GetPropertyStore(GETPROPERTYSTOREFLAGS flags, ref Guid riid, out IntPtr ppv);
        HRESULT GetPropertyStoreWithCreateObject(GETPROPERTYSTOREFLAGS flags, IntPtr punkCreateObject, ref Guid riid, out IntPtr ppv);
        HRESULT GetPropertyStoreForKeys([In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] PROPERTYKEY[] rgKeys, uint cKeys,
           GETPROPERTYSTOREFLAGS flags, ref Guid riid, out IntPtr ppv);
        HRESULT GetPropertyDescriptionList(PROPERTYKEY keyType, ref Guid riid, out IntPtr ppv);
        HRESULT Update(IBindCtx pbc);
        HRESULT GetProperty(PROPERTYKEY key, out PROPVARIANT ppropvar);
        HRESULT GetCLSID(PROPERTYKEY key, out Guid pclsid);
        HRESULT GetFileTime(PROPERTYKEY key, out FILETIME pft);
        HRESULT GetInt32(PROPERTYKEY key, out int pi);
        [PreserveSig()]
        HRESULT GetString(PROPERTYKEY key, out IntPtr ppsz);
        HRESULT Getuint32(PROPERTYKEY key, out uint pui);
        HRESULT Getuint64(PROPERTYKEY key, out ulong pull);
        HRESULT GetBool(PROPERTYKEY key, out bool pf);
    }

    public enum SIGDN : int
    {
        SIGDN_NORMALDISPLAY = 0x0,
        SIGDN_PARENTRELATIVEPARSING = unchecked((int)0x80018001),
        SIGDN_DESKTOPABSOLUTEPARSING = unchecked((int)0x80028000),
        SIGDN_PARENTRELATIVEEDITING = unchecked((int)0x80031001),
        SIGDN_DESKTOPABSOLUTEEDITING = unchecked((int)0x8004C000),
        SIGDN_FILESYSPATH = unchecked((int)0x80058000),
        SIGDN_URL = unchecked((int)0x80068000),
        SIGDN_PARENTRELATIVEFORADDRESSBAR = unchecked((int)0x8007C001),
        SIGDN_PARENTRELATIVE = unchecked((int)0x80080001)
    }

    public enum GETPROPERTYSTOREFLAGS
    {
        GPS_DEFAULT = 0,
        GPS_HANDLERPROPERTIESONLY = 0x1,
        GPS_READWRITE = 0x2,
        GPS_TEMPORARY = 0x4,
        GPS_FASTPROPERTIESONLY = 0x8,
        GPS_OPENSLOWITEM = 0x10,
        GPS_DELAYCREATION = 0x20,
        GPS_BESTEFFORT = 0x40,
        GPS_NO_OPLOCK = 0x80,
        GPS_PREFERQUERYPROPERTIES = 0x100,
        GPS_EXTRINSICPROPERTIES = 0x200,
        GPS_EXTRINSICPROPERTIESONLY = 0x400,
        GPS_VOLATILEPROPERTIES = 0x800,
        GPS_VOLATILEPROPERTIESONLY = 0x1000,
        GPS_MASK_VALID = 0x1fff
    }

    [ComImport()]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    public interface IShellLink
    {
        HRESULT GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, ref WIN32_FIND_DATA pfd, SLGP_FLAGS fFlags);
        HRESULT GetIDList(out IntPtr ppidl);
        HRESULT SetIDList(IntPtr pidl);
        HRESULT GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
        HRESULT SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        HRESULT GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
        HRESULT SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        HRESULT GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
        HRESULT SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        HRESULT GetHotKey(out ushort pwHotkey);
        HRESULT SetHotKey(ushort wHotKey);
        HRESULT GetShowCmd(out int piShowCmd);
        HRESULT SetShowCmd(int iShowCmd);
        HRESULT GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
        HRESULT SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        HRESULT SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        HRESULT Resolve(IntPtr hwnd, SLR_FLAGS fFlags);
        HRESULT SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    public enum SLGP_FLAGS
    {
        SLGP_SHORTPATH = 0x1,
        SLGP_UNCPRIORITY = 0x2,
        SLGP_RAWPATH = 0x4,
        SLGP_RELATIVEPRIORITY = 0x8
    }

    public enum SLR_FLAGS
    {
        SLR_NONE = 0,
        SLR_NO_UI = 0x1,
        SLR_ANY_MATCH = 0x2,
        SLR_UPDATE = 0x4,
        SLR_NOUPDATE = 0x8,
        SLR_NOSEARCH = 0x10,
        SLR_NOTRACK = 0x20,
        SLR_NOLINKINFO = 0x40,
        SLR_INVOKE_MSI = 0x80,
        SLR_NO_UI_WITH_MSG_PUMP = 0x101,
        SLR_OFFER_DELETE_WITHOUT_FILE = 0x200,
        SLR_KNOWNFOLDER = 0x400,
        SLR_MACHINE_IN_LOCAL_TARGET = 0x800,
        SLR_UPDATE_MACHINE_AND_SID = 0x1000,
        SLR_NO_OBJECT_ID = 0x2000
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4, CharSet = CharSet.Unicode)]
    public struct WIN32_FIND_DATA
    {
        public uint dwFileAttributes;
        public FILETIME ftCreationTime;
        public FILETIME ftLastAccessTime;
        public FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
        public string cAlternateFileName;
    }

    [ComImport]
    [Guid("000214E6-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellFolder
    {
        HRESULT ParseDisplayName(IntPtr hwnd,
            // IBindCtx pbc,
            IntPtr pbc,
            [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, [In, Out] ref uint pchEaten, out IntPtr ppidl, [In, Out] ref SFGAO pdwAttributes);
        HRESULT EnumObjects(IntPtr hwnd, SHCONTF grfFlags, out IEnumIDList ppenumIDList);
        [PreserveSig()]
        HRESULT BindToObject(IntPtr pidl,
            //IBindCtx pbc,
            IntPtr pbc,
            [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
        HRESULT BindToStorage(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
        HRESULT CompareIDs(IntPtr lParam, IntPtr pidl1, IntPtr pidl2);
        //HRESULT CreateViewObject(IntPtr hwndOwner, [In] ref Guid riid, out IntPtr ppv);
        HRESULT CreateViewObject(IntPtr hwndOwner, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
        HRESULT GetAttributesOf(uint cidl, IntPtr apidl, [In, Out] ref SFGAO rgfInOut);
        [PreserveSig()]
        HRESULT GetUIObjectOf(IntPtr hwndOwner, uint cidl, ref IntPtr apidl, [In] ref Guid riid, [In, Out] ref uint rgfReserved, out IntPtr ppv);
        HRESULT GetDisplayNameOf(IntPtr pidl, SHGDNF uFlags, out STRRET pName);
        HRESULT SetNameOf(IntPtr hwnd, IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszName, SHGDNF uFlags, out IntPtr ppidlOut);
    }

    [Flags]
    public enum SFGAO : uint
    {
        CANCOPY = 0x00000001,
        CANMOVE = 0x00000002,
        CANLINK = 0x00000004,
        STORAGE = 0x00000008,
        CANRENAME = 0x00000010,
        CANDELETE = 0x00000020,
        HASPROPSHEET = 0x00000040,
        DROPTARGET = 0x00000100,
        CAPABILITYMASK = 0x00000177,
        ENCRYPTED = 0x00002000,
        ISSLOW = 0x00004000,
        GHOSTED = 0x00008000,
        LINK = 0x00010000,
        SHARE = 0x00020000,
        READONLY = 0x00040000,
        HIDDEN = 0x00080000,
        DISPLAYATTRMASK = 0x000FC000,
        STREAM = 0x00400000,
        STORAGEANCESTOR = 0x00800000,
        VALIDATE = 0x01000000,
        REMOVABLE = 0x02000000,
        COMPRESSED = 0x04000000,
        BROWSABLE = 0x08000000,
        FILESYSANCESTOR = 0x10000000,
        FOLDER = 0x20000000,
        FILESYSTEM = 0x40000000,
        HASSUBFOLDER = 0x80000000,
        CONTENTSMASK = 0x80000000,
        STORAGECAPMASK = 0x70C50008,
        PKEYSFGAOMASK = 0x81044000
    }

    public enum SHGDNF
    {
        SHGDN_NORMAL = 0,
        SHGDN_INFOLDER = 0x1,
        SHGDN_FOREDITING = 0x1000,
        SHGDN_FORADDRESSBAR = 0x4000,
        SHGDN_FORPARSING = 0x8000
    }

    [StructLayout(LayoutKind.Explicit, Size = 264)]
    public struct STRRET
    {
        [FieldOffset(0)]
        public uint uType;
        [FieldOffset(4)]
        public IntPtr pOleStr;
        [FieldOffset(4)]
        public uint uOffset;
        [FieldOffset(4)]
        public IntPtr cStr;
    }

    [ComImport]
    [Guid("000214F2-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IEnumIDList
    {
        [PreserveSig()]
        HRESULT Next(uint celt, out IntPtr rgelt, out int pceltFetched);
        [PreserveSig()]
        HRESULT Skip(uint celt);
        void Reset();
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumIDList Clone();
    }

    [Flags]
    public enum SHCONTF : ushort
    {
        SHCONTF_CHECKING_FOR_CHILDREN = 0x0010,
        SHCONTF_FOLDERS = 0x0020,
        SHCONTF_NONFOLDERS = 0x0040,
        SHCONTF_INCLUDEHIDDEN = 0x0080,
        SHCONTF_INIT_ON_FIRST_NEXT = 0x0100,
        SHCONTF_NETPRINTERSRCH = 0x0200,
        SHCONTF_SHAREABLE = 0x0400,
        SHCONTF_STORAGE = 0x0800,
        SHCONTF_NAVIGATION_ENUM = 0x1000,
        SHCONTF_FASTITEMS = 0x2000,
        SHCONTF_FLATLIST = 0x4000,
        SHCONTF_ENABLE_ASYNC = 0x8000
    }

    [ComImport]
    [Guid("000214eb-0000-0000-c000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IExtractIcon
    {
        HRESULT GetIconLocation(uint uFlags, out string pszIconFile, uint cchMax, out int piIndex, out uint pwFlags);
        HRESULT Extract(string pszFile, uint nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, uint nIconSize);
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SHDESCRIPTIONID
    {
        public uint dwDescriptionId;
        public Guid clsid;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SHChangeNotifyEntry
    {
        public IntPtr pidl;
        public bool fRecursive;
    }
  
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 1)]
    public struct SHChangeDWORDAsIDList
    {
        public ushort cb;
        public uint dwItem1;
        public uint dwItem2;
        public ushort cbZero;
    }
}