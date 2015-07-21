using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices.ComTypes;

namespace Excellizer.Control
{
    public delegate void ProcessUrlActionEventHandler(object sender, ProcessUrlActionEventArgs e);


    public delegate void DocumentCompleteEventHandler(object sender, DocumentCompleteEventArgs e);
    public class DocumentCompleteEventArgs : System.EventArgs
    {
        public DocumentCompleteEventArgs() { }
        public void SetParameters(object Browser, string Url, bool IsTopLevel)
        {
            this.browser = Browser;
            this.url = Url;
            this.istoplevel = IsTopLevel;
        }
        public void Reset()
        {
            this.browser = null;
            this.url = string.Empty;
            this.istoplevel = false;
        }

        public object browser;
        public string url;
        public bool istoplevel;

    }

    public sealed class Iid_Clsids
    {
        //SID_STopWindow = {49e1b500-4636-11d3-97f7-00c04f45d0b3}
        public static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
        public static Guid IID_IViewObject = new Guid("0000010d-0000-0000-C000-000000000046");
        public static Guid IID_IAuthenticate = new Guid("79eac9d0-baf9-11ce-8c82-00aa004ba90b");
        public static Guid IID_IWindowForBindingUI = new Guid("79eac9d5-bafa-11ce-8c82-00aa004ba90b");
        public static Guid IID_IHttpSecurity = new Guid("79eac9d7-bafa-11ce-8c82-00aa004ba90b");
        //SID_SNewWindowManager same as IID_INewWindowManager
        public static Guid IID_INewWindowManager = new Guid("D2BC4C84-3F72-4a52-A604-7BCBF3982CBB");
        public static Guid IID_IOleClientSite = new Guid("00000118-0000-0000-C000-000000000046");
        public static Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
        public static Guid IID_TopLevelBrowser = new Guid("4C96BE40-915C-11CF-99D3-00AA004AE837");
        public static Guid IID_WebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
        public static Guid IID_IBinding = new Guid("79EAC9C0-BAF9-11CE-8C82-00AA004BA90B");
        public static Guid IID_IBindStatusCallBack = new Guid("79EAC9C1-BAF9-11CE-8C82-00AA004BA90B");
        public static Guid IID_IOleObject = new Guid("00000112-0000-0000-C000-000000000046");
        public static Guid IID_IOleWindow = new Guid("00000114-0000-0000-C000-000000000046");
        public static Guid IID_IServiceProvider = new Guid("6d5140c1-7436-11ce-8034-00aa006009fa");
        public static Guid IID_IWebBrowser = new Guid("EAB22AC1-30C1-11CF-A7EB-0000C05BAE0B");
        public static Guid IID_IWebBrowser2 = new Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E");
        public static Guid CLSID_WebBrowser = new Guid("8856F961-340A-11D0-A96B-00C04FD705A2");
        public static Guid CLSID_CGI_IWebBrowser = new Guid("ED016940-BD5B-11CF-BA4E-00C04FD70816");
        public static Guid CLSID_CGID_DocHostCommandHandler = new Guid("F38BC242-B950-11D1-8918-00C04FC2C836");
        public static Guid CLSID_ShellUIHelper = new Guid("64AB4BB7-111E-11D1-8F79-00C04FC2FBE1");
        public static Guid CLSID_SecurityManager = new Guid("7b8a2d94-0ac9-11d1-896c-00c04fb6bfc4");
        public static Guid IID_IShellUIHelper = new Guid("729FE2F8-1EA8-11d1-8F85-00C04FC2FBE1");
        public static Guid Guid_MSHTML = new Guid("DE4BA900-59CA-11CF-9592-444553540000");
        public static Guid CLSID_InternetSecurityManager = new Guid("7b8a2d94-0ac9-11d1-896c-00c04fB6bfc4");
        public static Guid IID_IInternetSecurityManager = new Guid("79EAC9EE-BAF9-11CE-8C82-00AA004BA90B");
        public static Guid CLSID_InternetZoneManager = new Guid("7B8A2D95-0AC9-11D1-896C-00C04FB6BFC4");
        public static Guid CGID_ShellDocView = new Guid("000214D1-0000-0000-C000-000000000046");
        //SID_SDownloadManager same as IID
        public static Guid SID_SDownloadManager = new Guid("988934A4-064B-11D3-BB80-00104B35E7F9");
        public static Guid IID_IDownloadManager = new Guid("988934A4-064B-11D3-BB80-00104B35E7F9");
        public static Guid IID_IHttpNegotiate = new Guid("79eac9d2-baf9-11ce-8c82-00aa004ba90b");
        public static Guid IID_IStream = new Guid("0000000c-0000-0000-C000-000000000046");
        public static Guid DIID_HTMLDocumentEvents2 = new Guid("3050f613-98b5-11cf-bb82-00aa00bdce0b");
        public static Guid DIID_HTMLWindowEvents2 = new Guid("3050f625-98b5-11cf-bb82-00aa00bdce0b");
        public static Guid DIID_HTMLElementEvents2 = new Guid("3050f60f-98b5-11cf-bb82-00aa00bdce0b");

        public static Guid IID_IDataObject = new Guid("0000010e-0000-0000-C000-000000000046");

        public static Guid CLSID_InternetShortcut = new Guid("FBF23B40-E3F0-101B-8488-00AA003E56F8");
        public static Guid IID_IUniformResourceLocatorA = new Guid("FBF23B80-E3F0-101B-8488-00AA003E56F8");
        public static Guid IID_IUniformResourceLocatorW = new Guid("CABB0DA0-DA57-11CF-9974-0020AFD79762");
        public static Guid IID_IHTMLBodyElement = new Guid("3050F1D8-98B5-11CF-BB82-00AA00BDCE0B");

        public static Guid CLSID_CUrlHistory = new Guid("3C374A40-BAE4-11CF-BF7D-00AA006946EE");

        public static Guid CLSID_HTMLDocument = new Guid("25336920-03F9-11cf-8FD0-00AA00686F13");
        public static Guid IID_IPropertyNotifySink = new Guid("9BFBBC02-EFF1-101A-84ED-00AA00341D07");

        public static Guid IID_IProtectFocus = new Guid("D81F90A3-8156-44F7-AD28-5ABB87003274");

        public static Guid IID_IHTMLOMWindowServices = new Guid("3050f5fc-98b5-11cf-bb82-00aa00bdce0b");
    }

    public sealed class Hresults
    {
        public const int NOERROR = 0;
        public const int S_OK = 0;
        public const int S_FALSE = 1;
        public const int E_PENDING = unchecked((int)0x8000000A);
        public const int E_HANDLE = unchecked((int)0x80070006);
        public const int E_NOTIMPL = unchecked((int)0x80004001);
        public const int E_NOINTERFACE = unchecked((int)0x80004002);
        //ArgumentNullException. NullReferenceException uses COR_E_NULLREFERENCE
        public const int E_POINTER = unchecked((int)0x80004003);
        public const int E_ABORT = unchecked((int)0x80004004);
        public const int E_FAIL = unchecked((int)0x80004005);
        public const int E_OUTOFMEMORY = unchecked((int)0x8007000E);
        public const int E_ACCESSDENIED = unchecked((int)0x80070005);
        public const int E_UNEXPECTED = unchecked((int)0x8000FFFF);
        public const int E_FLAGS = unchecked((int)0x1000);
        public const int E_INVALIDARG = unchecked((int)0x80070057);

        //Wininet
        public const int ERROR_SUCCESS = 0;
        public const int ERROR_FILE_NOT_FOUND = 2;
        public const int ERROR_ACCESS_DENIED = 5;
        public const int ERROR_INSUFFICIENT_BUFFER = 122;
        public const int ERROR_NO_MORE_ITEMS = 259;

        //Ole Errors
        public const int OLE_E_FIRST = unchecked((int)0x80040000);
        public const int OLE_E_LAST = unchecked((int)0x800400FF);
        public const int OLE_S_FIRST = unchecked((int)0x00040000);
        public const int OLE_S_LAST = unchecked((int)0x000400FF);
        //OLECMDERR_E_FIRST = 0x80040100
        public const int OLECMDERR_E_FIRST = unchecked((int)(OLE_E_LAST + 1));
        public const int OLECMDERR_E_NOTSUPPORTED = unchecked((int)(OLECMDERR_E_FIRST));
        public const int OLECMDERR_E_DISABLED = unchecked((int)(OLECMDERR_E_FIRST + 1));
        public const int OLECMDERR_E_NOHELP = unchecked((int)(OLECMDERR_E_FIRST + 2));
        public const int OLECMDERR_E_CANCELED = unchecked((int)(OLECMDERR_E_FIRST + 3));
        public const int OLECMDERR_E_UNKNOWNGROUP = unchecked((int)(OLECMDERR_E_FIRST + 4));

        public const int OLEOBJ_E_NOVERBS = unchecked((int)0x80040180);
        public const int OLEOBJ_S_INVALIDVERB = unchecked((int)0x00040180);
        public const int OLEOBJ_S_CANNOT_DOVERB_NOW = unchecked((int)0x00040181);
        public const int OLEOBJ_S_INVALIDHWND = unchecked((int)0x00040182);

        public const int DV_E_LINDEX = unchecked((int)0x80040068);
        public const int OLE_E_OLEVERB = unchecked((int)0x80040000);
        public const int OLE_E_ADVF = unchecked((int)0x80040001);
        public const int OLE_E_ENUM_NOMORE = unchecked((int)0x80040002);
        public const int OLE_E_ADVISENOTSUPPORTED = unchecked((int)0x80040003);
        public const int OLE_E_NOCONNECTION = unchecked((int)0x80040004);
        public const int OLE_E_NOTRUNNING = unchecked((int)0x80040005);
        public const int OLE_E_NOCACHE = unchecked((int)0x80040006);
        public const int OLE_E_BLANK = unchecked((int)0x80040007);
        public const int OLE_E_CLASSDIFF = unchecked((int)0x80040008);
        public const int OLE_E_CANT_GETMONIKER = unchecked((int)0x80040009);
        public const int OLE_E_CANT_BINDTOSOURCE = unchecked((int)0x8004000A);
        public const int OLE_E_STATIC = unchecked((int)0x8004000B);
        public const int OLE_E_PROMPTSAVECANCELLED = unchecked((int)0x8004000C);
        public const int OLE_E_INVALIDRECT = unchecked((int)0x8004000D);
        public const int OLE_E_WRONGCOMPOBJ = unchecked((int)0x8004000E);
        public const int OLE_E_INVALIDHWND = unchecked((int)0x8004000F);
        public const int OLE_E_NOT_INPLACEACTIVE = unchecked((int)0x80040010);
        public const int OLE_E_CANTCONVERT = unchecked((int)0x80040011);
        public const int OLE_E_NOSTORAGE = unchecked((int)0x80040012);
        public const int RPC_E_RETRY = unchecked((int)0x80010109);
    }

    public enum ProcessUrlActionFlags : uint
    {
        PUAF_DEFAULT = 0,
        PUAF_NOUI = 0x1,
        PUAF_ISFILE = 0x2,
        PUAF_WARN_IF_DENIED = 0x4,
        PUAF_FORCEUI_FOREGROUND = 0x8,
        PUAF_CHECK_TIFS = 0x10,
        PUAF_DONTCHECKBOXINDIALOG = 0x20,
        PUAF_TRUSTED = 0x40,
        PUAF_ACCEPT_WILDCARD_SCHEME = 0x80,
        PUAF_ENFORCERESTRICTED = 0x100
    }

    public enum URLPOLICY : uint
    {
        // Permissions 
        ALLOW = 0x00,
        QUERY = 0x01,
        DISALLOW = 0x03,

        ACTIVEX_CHECK_LIST = 0x00010000,
        CREDENTIALS_SILENT_LOGON_OK = 0x00000000,
        CREDENTIALS_MUST_PROMPT_USER = 0x00010000,
        CREDENTIALS_CONDITIONAL_PROMPT = 0x00020000,
        CREDENTIALS_ANONYMOUS_ONLY = 0x00030000,
        AUTHENTICATE_CLEARTEXT_OK = 0x00000000,
        AUTHENTICATE_CHALLENGE_RESPONSE = 0x00010000,
        AUTHENTICATE_MUTUAL_ONLY = 0x00030000,
        JAVA_PROHIBIT = 0x00000000,
        JAVA_HIGH = 0x00010000,
        JAVA_MEDIUM = 0x00020000,
        JAVA_LOW = 0x00030000,
        JAVA_CUSTOM = 0x00800000,
        CHANNEL_SOFTDIST_PROHIBIT = 0x00010000,
        CHANNEL_SOFTDIST_PRECACHE = 0x00020000,
        CHANNEL_SOFTDIST_AUTOINSTALL = 0x00030000,

        // For each action specified above the system maintains
        // a set of policies for the action. 
        // The only policies supported currently are permissions (i.e. is something allowed)
        // and logging status. 
        // IMPORTANT: If you are defining your own policies don't overload the meaning of the
        // loword of the policy. You can use the hiword to store any policy bits which are only
        // meaningful to your action.
        // For an example of how to do this look at the URLPOLICY_JAVA above

        // Notifications are not done when user already queried.
        NOTIFY_ON_ALLOW = 0x10,
        NOTIFY_ON_DISALLOW = 0x20,

        // Logging is done regardless of whether user was queried.
        LOG_ON_ALLOW = 0x40,
        LOG_ON_DISALLOW = 0x80,
        DONTCHECKDLGBOX = 0x100
    }

    public class ProcessUrlActionEventArgs : System.ComponentModel.CancelEventArgs
    {
        public bool handled;
        public bool hasContext;
        public string url;
        public URLACTION urlAction;
        public URLPOLICY urlPolicy;
        public Guid context;
        public ProcessUrlActionFlags flags;

        public ProcessUrlActionEventArgs() { }

        public void SetParameters(string surl, URLACTION action, URLPOLICY policy, Guid gcontext, ProcessUrlActionFlags puaf, bool bhascontext)
        {
            this.Cancel = false;
            this.handled = false;

            this.url = surl;
            this.urlAction = action;
            this.urlPolicy = policy;
            this.context = gcontext;
            this.flags = puaf;
            this.hasContext = bhascontext;
        }

        public void ResetParameters()
        {
            this.Cancel = false;
            this.handled = false;
            this.url = string.Empty;
            this.urlAction = URLACTION.MIN;
            this.urlPolicy = URLPOLICY.ALLOW;
            this.context = Guid.Empty;
            this.flags = ProcessUrlActionFlags.PUAF_DEFAULT;
            this.hasContext = false;
        }
    }

    public enum URLACTION : uint
    {
        // The zone manager maintains policies for a set of standard actions. 
        // These actions are identified by integral values (called action indexes)
        // specified below.

        // Minimum legal value for an action    
        MIN = 0x00001000,

        DOWNLOAD_MIN = 0x00001000,
        DOWNLOAD_SIGNED_ACTIVEX = 0x00001001,
        DOWNLOAD_UNSIGNED_ACTIVEX = 0x00001004,
        DOWNLOAD_CURR_MAX = 0x00001004,
        DOWNLOAD_MAX = 0x000011FF,

        ACTIVEX_MIN = 0x00001200,
        ACTIVEX_RUN = 0x00001200,
        ACTIVEX_OVERRIDE_OBJECT_SAFETY = 0x00001201, // aggregate next four
        ACTIVEX_OVERRIDE_DATA_SAFETY = 0x00001202, //
        ACTIVEX_OVERRIDE_SCRIPT_SAFETY = 0x00001203, //
        SCRIPT_OVERRIDE_SAFETY = 0x00001401, //
        ACTIVEX_CONFIRM_NOOBJECTSAFETY = 0x00001204, //
        ACTIVEX_TREATASUNTRUSTED = 0x00001205,
        ACTIVEX_NO_WEBOC_SCRIPT = 0x00001206,
        ACTIVEX_CURR_MAX = 0x00001206,
        ACTIVEX_MAX = 0x000013ff,

        SCRIPT_MIN = 0x00001400,
        SCRIPT_RUN = 0x00001400,
        SCRIPT_JAVA_USE = 0x00001402,
        SCRIPT_SAFE_ACTIVEX = 0x00001405,
        CROSS_DOMAIN_DATA = 0x00001406,
        SCRIPT_PASTE = 0x00001407,
        SCRIPT_CURR_MAX = 0x00001407,
        SCRIPT_MAX = 0x000015ff,

        HTML_MIN = 0x00001600,
        HTML_SUBMIT_FORMS = 0x00001601, // aggregate next two
        HTML_SUBMIT_FORMS_FROM = 0x00001602, //
        HTML_SUBMIT_FORMS_TO = 0x00001603, //
        HTML_FONT_DOWNLOAD = 0x00001604,
        HTML_JAVA_RUN = 0x00001605, // derive from Java custom policy
        HTML_USERDATA_SAVE = 0x00001606,
        HTML_SUBFRAME_NAVIGATE = 0x00001607,
        HTML_META_REFRESH = 0x00001608,
        HTML_MIXED_CONTENT = 0x00001609,
        HTML_MAX = 0x000017ff,

        SHELL_MIN = 0x00001800,
        SHELL_INSTALL_DTITEMS = 0x00001800,
        SHELL_MOVE_OR_COPY = 0x00001802,
        SHELL_FILE_DOWNLOAD = 0x00001803,
        SHELL_VERB = 0x00001804,
        SHELL_WEBVIEW_VERB = 0x00001805,
        SHELL_SHELLEXECUTE = 0x00001806,
        SHELL_CURR_MAX = 0x00001806,
        SHELL_MAX = 0x000019ff,

        NETWORK_MIN = 0x00001A00,
        CREDENTIALS_USE = 0x00001A00,
        AUTHENTICATE_CLIENT = 0x00001A01,

        COOKIES = 0x00001A02,
        COOKIES_SESSION = 0x00001A03,
        CLIENT_CERT_PROMPT = 0x00001A0,
        COOKIES_THIRD_PARTY = 0x00001A05,
        COOKIES_SESSION_THIRD_PARTY = 0x00001A06,
        COOKIES_ENABLED = 0x00001A10,
        NETWORK_CURR_MAX = 0x00001A10,
        NETWORK_MAX = 0x00001Bff,

        JAVA_MIN = 0x00001C00,
        JAVA_PERMISSIONS = 0x00001C00,
        JAVA_CURR_MAX = 0x00001C00,
        JAVA_MAX = 0x00001Cff,

        // The following Infodelivery actions should have no default policies
        // in the registry.  They assume that no default policy means fall
        // back to the global restriction.  If an admin sets a policy per
        // zone, then it overrides the global restriction.

        INFODELIVERY_MIN = 0x00001D00,
        INFODELIVERY_NO_ADDING_CHANNELS = 0x00001D00,
        INFODELIVERY_NO_EDITING_CHANNELS = 0x00001D01,
        INFODELIVERY_NO_REMOVING_CHANNELS = 0x00001D02,
        INFODELIVERY_NO_ADDING_SUBSCRIPTIONS = 0x00001D03,
        INFODELIVERY_NO_EDITING_SUBSCRIPTIONS = 0x00001D04,
        INFODELIVERY_NO_REMOVING_SUBSCRIPTIONS = 0x00001D05,
        INFODELIVERY_NO_CHANNEL_LOGGING = 0x00001D06,
        INFODELIVERY_CURR_MAX = 0x00001D06,
        INFODELIVERY_MAX = 0x00001Dff,
        CHANNEL_SOFTDIST_MIN = 0x00001E00,
        CHANNEL_SOFTDIST_PERMISSIONS = 0x00001E05,
        CHANNEL_SOFTDIST_MAX = 0x00001Eff
    }

    #region IServiceProvider Interface
    //MIDL_INTERFACE("6d5140c1-7436-11ce-8034-00aa006009fa")
    //IServiceProvider : public IUnknown
    //{
    //public:
    //    virtual /* [local] */ HRESULT STDMETHODCALLTYPE QueryService( 
    //        /* [in] */ REFGUID guidService,
    //        /* [in] */ REFIID riid,
    //        /* [out] */ void __RPC_FAR *__RPC_FAR *ppvObject) = 0;

    //    template <class Q>
    //    HRESULT STDMETHODCALLTYPE QueryService(REFGUID guidService, Q** pp)
    //    {
    //        return QueryService(guidService, __uuidof(Q), (void **)pp);
    //    }
    //};
    [ComImport, ComVisible(true)]
    [Guid("6d5140c1-7436-11ce-8034-00aa006009fa")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IServiceProvider
    {
        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int QueryService(
            [In] ref Guid guidService,
            [In] ref Guid riid,
            [Out] out IntPtr ppvObject);
        //This does not work i.e.-> ppvObject = (INewWindowManager)this
        //[Out, MarshalAs(UnmanagedType.Interface)] out object ppvObject);
    }
    #endregion
    
    /*
    [ComImport, GuidAttribute("6D5140C1-7436-11CE-8034-00AA006009FA")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IServiceProvider
    {
        void QueryService(ref Guid guidService, ref Guid riid,
                  [MarshalAs(UnmanagedType.Interface)] out object ppvObject);
    }
    */

    #region IInternetSecurityManager Interface
    [ComVisible(true), ComImport,
    GuidAttribute("79EAC9EE-BAF9-11CE-8C82-00AA004BA90B"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInternetSecurityManager
    {
        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int SetSecuritySite(
            [In] IntPtr pSite);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int GetSecuritySite(
            out IntPtr pSite);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int MapUrlToZone(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
            out UInt32 pdwZone,
            [In] UInt32 dwFlags);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int GetSecurityId(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
            [Out] IntPtr pbSecurityId, [In, Out] ref UInt32 pcbSecurityId,
            [In] ref UInt32 dwReserved);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int ProcessUrlAction(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
            UInt32 dwAction,
            IntPtr pPolicy, UInt32 cbPolicy,
            IntPtr pContext, UInt32 cbContext,
            UInt32 dwFlags,
            UInt32 dwReserved);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int QueryCustomPolicy(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
            ref Guid guidKey,
            out IntPtr ppPolicy, out UInt32 pcbPolicy,
            IntPtr pContext, UInt32 cbContext,
            UInt32 dwReserved);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int SetZoneMapping(
            UInt32 dwZone,
            [In, MarshalAs(UnmanagedType.LPWStr)] string lpszPattern,
            UInt32 dwFlags);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int GetZoneMappings(
            [In] UInt32 dwZone,
            out IEnumString ppenumString,
            [In] UInt32 dwFlags);
    }
    #endregion
    /*
    [ComImport, GuidAttribute("79EAC9EE-BAF9-11CE-8C82-00AA004BA90B")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInternetSecurityManager
    {
        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int SetSecuritySite([In] IntPtr pSite);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int GetSecuritySite([Out] IntPtr pSite);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int MapUrlToZone([In, MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
                 out UInt32 pdwZone, UInt32 dwFlags);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int GetSecurityId([MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
                  [MarshalAs(UnmanagedType.LPArray)] byte[] pbSecurityId,
                  ref UInt32 pcbSecurityId, uint dwReserved);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int ProcessUrlAction([In, MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
                 UInt32 dwAction, out byte pPolicy, UInt32 cbPolicy,
                 byte pContext, UInt32 cbContext, UInt32 dwFlags,
                 UInt32 dwReserved);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int QueryCustomPolicy([In, MarshalAs(UnmanagedType.LPWStr)] string pwszUrl,
                  ref Guid guidKey, ref byte ppPolicy, ref UInt32 pcbPolicy,
                  ref byte pContext, UInt32 cbContext, UInt32 dwReserved);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int SetZoneMapping(UInt32 dwZone,
                   [In, MarshalAs(UnmanagedType.LPWStr)] string lpszPattern,
                   UInt32 dwFlags);

        [return: MarshalAs(UnmanagedType.I4)]
        [PreserveSig]
        int GetZoneMappings(UInt32 dwZone, out UCOMIEnumString ppenumString,
                UInt32 dwFlags);
    }*/

    public enum WinInetErrors : int
    {
        HTTP_STATUS_CONTINUE = 100, //The request can be continued.
        HTTP_STATUS_SWITCH_PROTOCOLS = 101, //The server has switched protocols in an upgrade header.
        HTTP_STATUS_OK = 200, //The request completed successfully.
        HTTP_STATUS_CREATED = 201, //The request has been fulfilled and resulted in the creation of a new resource.
        HTTP_STATUS_ACCEPTED = 202, //The request has been accepted for processing, but the processing has not been completed.
        HTTP_STATUS_PARTIAL = 203, //The returned meta information in the entity-header is not the definitive set available from the origin server.
        HTTP_STATUS_NO_CONTENT = 204, //The server has fulfilled the request, but there is no new information to send back.
        HTTP_STATUS_RESET_CONTENT = 205, //The request has been completed, and the client program should reset the document view that caused the request to be sent to allow the user to easily initiate another input action.
        HTTP_STATUS_PARTIAL_CONTENT = 206, //The server has fulfilled the partial GET request for the resource.
        HTTP_STATUS_AMBIGUOUS = 300, //The server couldn't decide what to return.
        HTTP_STATUS_MOVED = 301, //The requested resource has been assigned to a new permanent URI (Uniform Resource Identifier), and any future references to this resource should be done using one of the returned URIs.
        HTTP_STATUS_REDIRECT = 302, //The requested resource resides temporarily under a different URI (Uniform Resource Identifier).
        HTTP_STATUS_REDIRECT_METHOD = 303, //The response to the request can be found under a different URI (Uniform Resource Identifier) and should be retrieved using a GET HTTP verb on that resource.
        HTTP_STATUS_NOT_MODIFIED = 304, //The requested resource has not been modified.
        HTTP_STATUS_USE_PROXY = 305, //The requested resource must be accessed through the proxy given by the location field.
        HTTP_STATUS_REDIRECT_KEEP_VERB = 307, //The redirected request keeps the same HTTP verb. HTTP/1.1 behavior.

        HTTP_STATUS_BAD_REQUEST = 400,
        HTTP_STATUS_DENIED = 401,
        HTTP_STATUS_PAYMENT_REQ = 402,
        HTTP_STATUS_FORBIDDEN = 403,
        HTTP_STATUS_NOT_FOUND = 404,
        HTTP_STATUS_BAD_METHOD = 405,
        HTTP_STATUS_NONE_ACCEPTABLE = 406,
        HTTP_STATUS_PROXY_AUTH_REQ = 407,
        HTTP_STATUS_REQUEST_TIMEOUT = 408,
        HTTP_STATUS_CONFLICT = 409,
        HTTP_STATUS_GONE = 410,
        HTTP_STATUS_LENGTH_REQUIRED = 411,
        HTTP_STATUS_PRECOND_FAILED = 412,
        HTTP_STATUS_REQUEST_TOO_LARGE = 413,
        HTTP_STATUS_URI_TOO_LONG = 414,
        HTTP_STATUS_UNSUPPORTED_MEDIA = 415,
        HTTP_STATUS_RETRY_WITH = 449,
        HTTP_STATUS_SERVER_ERROR = 500,
        HTTP_STATUS_NOT_SUPPORTED = 501,
        HTTP_STATUS_BAD_GATEWAY = 502,
        HTTP_STATUS_SERVICE_UNAVAIL = 503,
        HTTP_STATUS_GATEWAY_TIMEOUT = 504,
        HTTP_STATUS_VERSION_NOT_SUP = 505,

        ERROR_INTERNET_ASYNC_THREAD_FAILED = 12047,    //The application could not start an asynchronous thread.
        ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = 12166,    //There was an error in the automatic proxy configuration script.
        ERROR_INTERNET_BAD_OPTION_LENGTH = 12010,    //The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified.
        ERROR_INTERNET_BAD_REGISTRY_PARAMETER = 12022,    //A required registry value was located but is an incorrect type or has an invalid value.
        ERROR_INTERNET_CANNOT_CONNECT = 12029,    //The attempt to connect to the server failed.
        ERROR_INTERNET_CHG_POST_IS_NON_SECURE = 12042,    //The application is posting and attempting to change multiple lines of text on a server that is not secure.
        ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = 12044,    //The server is requesting client authentication.
        ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = 12046,    //Client authorization is not set up on this computer.
        ERROR_INTERNET_CONNECTION_ABORTED = 12030,    //The connection with the server has been terminated.
        ERROR_INTERNET_CONNECTION_RESET = 12031,    //The connection with the server has been reset.
        ERROR_INTERNET_DIALOG_PENDING = 12049,    //Another thread has a password dialog box in progress.
        ERROR_INTERNET_DISCONNECTED = 12163,    //The Internet connection has been lost.
        ERROR_INTERNET_EXTENDED_ERROR = 12003,    //An extended error was returned from the server. This is typically a string or buffer containing a verbose error message. Call InternetGetLastResponseInfo to retrieve the error text.
        ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = 12171,    //The function failed due to a security check.
        ERROR_INTERNET_FORCE_RETRY = 12032,    //The function needs to redo the request.
        ERROR_INTERNET_FORTEZZA_LOGIN_NEEDED = 12054,    //The requested resource requires Fortezza authentication.
        ERROR_INTERNET_HANDLE_EXISTS = 12036,    //The request failed because the handle already exists.
        ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = 12039,    //The application is moving from a non-SSL to an SSL connection because of a redirect.
        ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = 12052,    //The data being submitted to an SSL connection is being redirected to a non-SSL connection.
        ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = 12040,    //The application is moving from an SSL to an non-SSL connection because of a redirect.
        ERROR_INTERNET_INCORRECT_FORMAT = 12027,    //The format of the request is invalid.
        ERROR_INTERNET_INCORRECT_HANDLE_STATE = 12019,    //The requested operation cannot be carried out because the handle supplied is not in the correct state.
        ERROR_INTERNET_INCORRECT_HANDLE_TYPE = 12018,    //The type of handle supplied is incorrect for this operation.
        ERROR_INTERNET_INCORRECT_PASSWORD = 12014,    //The request to connect and log on to an FTP server could not be completed because the supplied password is incorrect.
        ERROR_INTERNET_INCORRECT_USER_NAME = 12013,    //The request to connect and log on to an FTP server could not be completed because the supplied user name is incorrect.
        ERROR_INTERNET_INSERT_CDROM = 12053,    //The request requires a CD-ROM to be inserted in the CD-ROM drive to locate the resource requested.
        ERROR_INTERNET_INTERNAL_ERROR = 12004,    //An internal error has occurred.
        ERROR_INTERNET_INVALID_CA = 12045,    //The function is unfamiliar with the Certificate Authority that generated the server's certificate.
        ERROR_INTERNET_INVALID_OPERATION = 12016,    //The requested operation is invalid.
        ERROR_INTERNET_INVALID_OPTION = 12009,    //A request to InternetQueryOption or InternetSetOption specified an invalid option value.
        ERROR_INTERNET_INVALID_PROXY_REQUEST = 12033,    //The request to the proxy was invalid.
        ERROR_INTERNET_INVALID_URL = 12005,    //The URL is invalid.
        ERROR_INTERNET_ITEM_NOT_FOUND = 12028,    //The requested item could not be located.
        ERROR_INTERNET_LOGIN_FAILURE = 12015,    //The request to connect and log on to an FTP server failed.
        ERROR_INTERNET_LOGIN_FAILURE_DISPLAY_ENTITY_BODY = 12174,    //The MS-Logoff digest header has been returned from the Web site. This header specifically instructs the digest package to purge credentials for the associated realm. This error will only be returned if INTERNET_ERROR_MASK_LOGIN_FAILURE_DISPLAY_ENTITY_BODY has been set.
        ERROR_INTERNET_MIXED_SECURITY = 12041,    //The content is not entirely secure. Some of the content being viewed may have come from unsecured servers.
        ERROR_INTERNET_NAME_NOT_RESOLVED = 12007,    //The server name could not be resolved.
        ERROR_INTERNET_NEED_MSN_SSPI_PKG = 12173,    //Not currently implemented.
        ERROR_INTERNET_NEED_UI = 12034,    //A user interface or other blocking operation has been requested.
        ERROR_INTERNET_NO_CALLBACK = 12025,    //An asynchronous request could not be made because a callback function has not been set.
        ERROR_INTERNET_NO_CONTEXT = 12024,    //An asynchronous request could not be made because a zero context value was supplied.
        ERROR_INTERNET_NO_DIRECT_ACCESS = 12023,    //Direct network access cannot be made at this time.
        ERROR_INTERNET_NOT_INITIALIZED = 12172,    //Initialization of the WinINet API has not occurred. Indicates that a higher-level function, such as InternetOpen, has not been called yet.
        ERROR_INTERNET_NOT_PROXY_REQUEST = 12020,    //The request cannot be made via a proxy.
        ERROR_INTERNET_OPERATION_CANCELLED = 12017,    //The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed.
        ERROR_INTERNET_OPTION_NOT_SETTABLE = 12011,    //The requested option cannot be set, only queried.
        ERROR_INTERNET_OUT_OF_HANDLES = 12001,    //No more handles could be generated at this time.
        ERROR_INTERNET_POST_IS_NON_SECURE = 12043,    //The application is posting data to a server that is not secure.
        ERROR_INTERNET_PROTOCOL_NOT_FOUND = 12008,    //The requested protocol could not be located.
        ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = 12165,    //The designated proxy server cannot be reached.
        ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = 12048,    //The function could not handle the redirection, because the scheme changed (for example, HTTP to FTP).
        ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = 12021,    //A required registry value could not be located.
        ERROR_INTERNET_REQUEST_PENDING = 12026,    //The required operation could not be completed because one or more requests are pending.
        ERROR_INTERNET_RETRY_DIALOG = 12050,    //The dialog box should be retried.
        ERROR_INTERNET_SEC_CERT_CN_INVALID = 12038,    //SSL certificate common name (host name field) is incorrect뾣or example, if you entered www.server.com and the common name on the certificate says www.different.com.
        ERROR_INTERNET_SEC_CERT_DATE_INVALID = 12037,    //SSL certificate date that was received from the server is bad. The certificate is expired.
        ERROR_INTERNET_SEC_CERT_ERRORS = 12055,    //The SSL certificate contains errors.
        ERROR_INTERNET_SEC_CERT_NO_REV = 12056,
        ERROR_INTERNET_SEC_CERT_REV_FAILED = 12057,
        ERROR_INTERNET_SEC_CERT_REVOKED = 12170,    //SSL certificate was revoked.
        ERROR_INTERNET_SEC_INVALID_CERT = 12169,    //SSL certificate is invalid.
        ERROR_INTERNET_SECURITY_CHANNEL_ERROR = 12157,    //The application experienced an internal error loading the SSL libraries.
        ERROR_INTERNET_SERVER_UNREACHABLE = 12164,    //The Web site or server indicated is unreachable.
        ERROR_INTERNET_SHUTDOWN = 12012,    //WinINet support is being shut down or unloaded.
        ERROR_INTERNET_TCPIP_NOT_INSTALLED = 12159,    //The required protocol stack is not loaded and the application cannot start WinSock.
        ERROR_INTERNET_TIMEOUT = 12002,    //The request has timed out.
        ERROR_INTERNET_UNABLE_TO_CACHE_FILE = 12158,    //The function was unable to cache the file.
        ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = 12167,    //The automatic proxy configuration script could not be downloaded. The INTERNET_FLAG_MUST_CACHE_REQUEST flag was set.

        INET_E_INVALID_URL = unchecked((int)0x800C0002),
        INET_E_NO_SESSION = unchecked((int)0x800C0003),
        INET_E_CANNOT_CONNECT = unchecked((int)0x800C0004),
        INET_E_RESOURCE_NOT_FOUND = unchecked((int)0x800C0005),
        INET_E_OBJECT_NOT_FOUND = unchecked((int)0x800C0006),
        INET_E_DATA_NOT_AVAILABLE = unchecked((int)0x800C0007),
        INET_E_DOWNLOAD_FAILURE = unchecked((int)0x800C0008),
        INET_E_AUTHENTICATION_REQUIRED = unchecked((int)0x800C0009),
        INET_E_NO_VALID_MEDIA = unchecked((int)0x800C000A),
        INET_E_CONNECTION_TIMEOUT = unchecked((int)0x800C000B),
        INET_E_DEFAULT_ACTION = unchecked((int)0x800C0011),
        INET_E_INVALID_REQUEST = unchecked((int)0x800C000C),
        INET_E_UNKNOWN_PROTOCOL = unchecked((int)0x800C000D),
        INET_E_QUERYOPTION_UNKNOWN = unchecked((int)0x800C0013),
        INET_E_SECURITY_PROBLEM = unchecked((int)0x800C000E),
        INET_E_CANNOT_LOAD_DATA = unchecked((int)0x800C000F),
        INET_E_CANNOT_INSTANTIATE_OBJECT = unchecked((int)0x800C0010),
        INET_E_REDIRECT_FAILED = unchecked((int)0x800C0014),
        INET_E_REDIRECT_TO_DIR = unchecked((int)0x800C0015),
        INET_E_CANNOT_LOCK_REQUEST = unchecked((int)0x800C0016),
        INET_E_USE_EXTEND_BINDING = unchecked((int)0x800C0017),
        INET_E_TERMINATED_BIND = unchecked((int)0x800C0018),
        INET_E_ERROR_FIRST = unchecked((int)0x800C0002),
        INET_E_CODE_DOWNLOAD_DECLINED = unchecked((int)0x800C0100),
        INET_E_RESULT_DISPATCHED = unchecked((int)0x800C0200),
        INET_E_CANNOT_REPLACE_SFP_FILE = unchecked((int)0x800C0300),

        HTTP_COOKIE_DECLINED = 12162,    //The HTTP cookie was declined by the server.
        HTTP_COOKIE_NEEDS_CONFIRMATION = 12161,    //The HTTP cookie requires confirmation.
        HTTP_DOWNLEVEL_SERVER = 12151,    //The server did not return any headers.
        HTTP_HEADER_ALREADY_EXISTS = 12155,    //The header could not be added because it already exists.
        HTTP_HEADER_NOT_FOUND = 12150,    //The requested header could not be located.
        HTTP_INVALID_HEADER = 12153,    //The supplied header is invalid.
        HTTP_INVALID_QUERY_REQUEST = 12154,    //The request made to HttpQueryInfo is invalid.
        HTTP_INVALID_SERVER_RESPONSE = 12152,    //The server response could not be parsed.
        HTTP_NOT_REDIRECTED = 12160,    //The HTTP request was not redirected.
        HTTP_REDIRECT_FAILED = 12156,    //The redirection failed because either the scheme changed (for example, HTTP to FTP) or all attempts made to redirect failed (default is five attempts).
        HTTP_REDIRECT_NEEDS_CONFIRMATION = 12168    //The redirection requires user confirmation.
    }

    /**

    //[Guid("<interface guid>")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _ZoneSecurityDemo
    {
        [DispId(1)]
        void AssessZoneSafety();
    }

    class ZoneSecurityDemo
    {
    }
    
    //[Guid("<class guid>")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("Excellizer.Control.ZoneSecurityDemo")]
    public class ZoneSecurityDemo : System.Windows.Forms.Control
    {
        private Guid _IID_TopLevelBrowser = new Guid("4C96BE40-915C-11CF-99D3-00AA004AE837");
        private Guid _IID_WebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
        private Guid _CLSID_SecurityManager = new Guid("7b8a2d94-0ac9-11d1-896c-00c04fb6bfc4");

        private bool _ZoneSafetyConfirmed = false;
        /*
        public void AssessZoneSafety()
        {
            object oleClientSiteObj = null;
            Excellizer.Control.IServiceProvider serviceProvider = null;
            object topServiceProviderObj = null;
            IServiceProvider topServiceProvider = null;
            object webBrowserObj = null;
            SHDocVw.IWebBrowser webBrowser = null;
            try
            {
                // Get the client site service provider.
                Type iOleObjectType = this.GetType().GetInterface("IOleObject", true);
                oleClientSiteObj = iOleObjectType.InvokeMember("GetClientSite",
                                           BindingFlags.Instance |
                                           BindingFlags.InvokeMethod |
                                           BindingFlags.Public, null,
                                           this, null);
                serviceProvider = oleClientSiteObj as Excellizer.Control.IServiceProvider;

                // Get top level browser service provider.
                Guid IID_TopLevelBrowser = _IID_TopLevelBrowser;
                Guid Riid = typeof(Excellizer.Control.IServiceProvider).GUID;
                topServiceProviderObj = null;
                serviceProvider.QueryService(ref IID_TopLevelBrowser, ref Riid,
                                 out topServiceProviderObj);
                topServiceProvider = topServiceProviderObj as IServiceProvider;

                // Get web browser object.
                Guid IID_WebBrowserApp = _IID_WebBrowserApp;
                Riid = typeof(SHDocVw.IWebBrowser).GUID;
                webBrowserObj = null;
                topServiceProvider.QueryService(ref IID_WebBrowserApp, ref Riid,
                                out webBrowserObj);
                webBrowser = webBrowserObj as SHDocVw.IWebBrowser;

                // Determine which zone the browser is currently in.
                Type t = Type.GetTypeFromCLSID(_CLSID_SecurityManager);
                object securityManager = Activator.CreateInstance(t);
                IInternetSecurityManager ISM = securityManager as IInternetSecurityManager;
                uint Zone;
                ISM.MapUrlToZone(webBrowser.LocationURL, out Zone, 0);
                Marshal.ReleaseComObject(securityManager);

                // Only accept calls from the My Computer zone.
                if (Zone == 0)
                    _ZoneSafetyConfirmed = true;
            }
            catch
            {
            }
            finally
            {
                if (webBrowser != null)
                    Marshal.ReleaseComObject(webBrowser);
                if (webBrowserObj != null)
                    Marshal.ReleaseComObject(webBrowserObj);
                if (topServiceProvider != null)
                    Marshal.ReleaseComObject(topServiceProvider);
                if (topServiceProviderObj != null)
                    Marshal.ReleaseComObject(topServiceProviderObj);
                if (serviceProvider != null)
                    Marshal.ReleaseComObject(serviceProvider);
                if (oleClientSiteObj != null)
                    Marshal.ReleaseComObject(oleClientSiteObj);
            }
        }
    }
     **/
}
