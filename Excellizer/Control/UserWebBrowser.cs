using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Excellizer.Control
{
    public partial class UserWebBrowser : UserControl, IInternetSecurityManager, IServiceProvider 
    {
        public event ProcessUrlActionEventHandler ProcessUrlAction = null;
        private ProcessUrlActionEventArgs ProcessUrlActionEvent = new ProcessUrlActionEventArgs();
        private IntPtr m_NullPointer = IntPtr.Zero;
        private bool m_UseInternalDownloadManager = true;
        ////private CSEXWBDLMANLib.csDLManClass m_csexwbCOMLib = null;

        public UserWebBrowser()
        {
            InitializeComponent();
            webBrowser.ScriptErrorsSuppressed = true;
        }
        /**
        private bool AddThisIEServerHwndToComLib()
        {
            if ((m_csexwbCOMLib.HWNDInternetExplorerServer == 0) &&
                (WBIEServerHandle() != IntPtr.Zero))
            {
                m_csexwbCOMLib.HWNDInternetExplorerServer = m_hWBServerHandle.ToInt32();
                return true;
            }
            else
                return false;
        }
        **/
        #region IServiceProvider Members

        int IServiceProvider.QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject)
        {
            int hr = Hresults.E_NOINTERFACE;
            ppvObject = m_NullPointer;

            if ((guidService == Iid_Clsids.SID_SDownloadManager) &&
                (riid == Iid_Clsids.IID_IDownloadManager) &&
                (m_UseInternalDownloadManager))
            {
                ////AddThisIEServerHwndToComLib();
                //QI for IDownloadManager interface from our COM object
                ////ppvObject = Marshal.GetComInterfaceForObject(m_csexwbCOMLib, typeof(IDownloadManager));
                hr = Hresults.S_OK;
            }
            else if (riid == Iid_Clsids.IID_IHttpSecurity)
            {
                ////ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IHttpSecurity));
                hr = Hresults.S_OK;

                //Ulternative
                //try
                //{
                //    m_pUnk = IntPtr.Zero;
                //    m_pTargetIface = IntPtr.Zero;
                //    m_pUnk = Marshal.GetIUnknownForObject(this);
                //    Marshal.QueryInterface(m_pUnk, ref IID_IHttpSecurity, out m_pTargetIface);
                //    Marshal.Release(m_pUnk);
                //    ppvObject = m_pTargetIface;
                //    hr = Hresults.S_OK;
                //}
                //catch (Exception)
                //{
                //}
            }
            else if (riid == Iid_Clsids.IID_INewWindowManager) //xpsp2
            {
                ////ppvObject = Marshal.GetComInterfaceForObject(this, typeof(INewWindowManager));
                hr = Hresults.S_OK;
            }
            else if (riid == Iid_Clsids.IID_IWindowForBindingUI)
            {
                ////ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IWindowForBindingUI));
                hr = Hresults.S_OK;
            }
            else if (guidService == Iid_Clsids.IID_IInternetSecurityManager)
            {
                ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IInternetSecurityManager));
                hr = Hresults.S_OK;
            }
            else if ((guidService == Iid_Clsids.IID_IAuthenticate)
               && (riid == Iid_Clsids.IID_IAuthenticate))
            {
                ////ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IAuthenticate));
                hr = Hresults.S_OK;
            }
            else if (riid == Iid_Clsids.IID_IProtectFocus) //IE7 + Vista
            {
                ////ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IProtectFocus));
                hr = Hresults.S_OK;
            }
            else if ((riid == Iid_Clsids.IID_IHTMLOMWindowServices) &&
                (guidService == Iid_Clsids.IID_IHTMLOMWindowServices))
            {
                ////ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IHTMLOMWindowServices));
                hr = Hresults.S_OK;
            }

            return hr;
        }

        #endregion

        #region IInternetSecurityManager Members

        int IInternetSecurityManager.SetSecuritySite(IntPtr pSite)
        {
            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        int IInternetSecurityManager.GetSecuritySite(out IntPtr pSite)
        {
            pSite = IntPtr.Zero;
            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        int IInternetSecurityManager.MapUrlToZone(string pwszUrl, out uint pdwZone, uint dwFlags)
        {
            // All URLs are on the local machine - most trusted and return S_OK;
            //pdwZone = (uint)tagURLZONE.URLZONE_LOCAL_MACHINE;
            //pdwZone =  (uint)tagURLZONE.URLZONE_INTRANET;
            //pdwZone =  (uint)tagURLZONE.URLZONE_TRUSTED; //....
            //return Hresults.S_OK;

            pdwZone = 1000;
            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        //private const string m_strSecurity = "None:localhost+My Computer";
        int IInternetSecurityManager.GetSecurityId(string pwszUrl, IntPtr pbSecurityId, ref uint pcbSecurityId, ref uint dwReserved)
        {
            //pbSecurityId = Marshal.StringToCoTaskMemAnsi(m_strSecurity);
            //pcbSecurityId = (uint)m_strSecurity.Length;
            //return Hresults.S_OK;
            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        /*
        MSDN:
        The current list of URLACTION that will not be passed to the custom security manager
        in most circumstances by Internet Explorer 5 are:
	        URLACTION_SHELL_FILE_DOWNLOAD 
	        URLACTION_COOKIES 
	        URLACTION_JAVA_PERMISSIONS 
	        URLACTION_SCRIPT_PASTE 
        There is no workaround for this problem. The behavior for the URLACTION can only be
        changed for all browser clients on the system by altering the security zone settings
        from Internet Options.
        */
        int IInternetSecurityManager.ProcessUrlAction(string pwszUrl, uint dwAction, IntPtr pPolicy, uint cbPolicy, IntPtr pContext, uint cbContext, uint dwFlags, uint dwReserved)
        {
            if (ProcessUrlAction == null)
                return (int)WinInetErrors.INET_E_DEFAULT_ACTION;

            try
            {
                URLACTION action = (URLACTION)dwAction;
                ProcessUrlActionFlags flags = (ProcessUrlActionFlags)dwFlags;
                bool hasUrlPolicy = (cbPolicy >= unchecked((uint)Marshal.SizeOf(typeof(int))));
                URLPOLICY urlPolicy = (hasUrlPolicy) ? urlPolicy = (URLPOLICY)Marshal.ReadInt32(pPolicy) : URLPOLICY.ALLOW;
                bool hasContext = (cbContext >= unchecked((uint)Marshal.SizeOf(typeof(Guid))));
                Guid context = (hasContext) ? (Guid)Marshal.PtrToStructure(pContext, typeof(Guid)) : Guid.Empty;

                ProcessUrlActionEvent.SetParameters(pwszUrl, action, urlPolicy, context, flags, hasContext);
                ProcessUrlAction(this, ProcessUrlActionEvent);

                if (ProcessUrlActionEvent.handled && hasUrlPolicy)
                {
                    Marshal.WriteInt32(pPolicy, (int)ProcessUrlActionEvent.urlPolicy);
                    return (ProcessUrlActionEvent.Cancel) ? Hresults.S_FALSE : Hresults.S_OK;
                }
            }
            finally
            {
                ProcessUrlActionEvent.ResetParameters();
            }

            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        int IInternetSecurityManager.QueryCustomPolicy(string pwszUrl, ref Guid guidKey, out IntPtr ppPolicy, out uint pcbPolicy, IntPtr pContext, uint cbContext, uint dwReserved)
        {
            ppPolicy = IntPtr.Zero;
            pcbPolicy = 0;
            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        int IInternetSecurityManager.SetZoneMapping(uint dwZone, string lpszPattern, uint dwFlags)
        {
            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        int IInternetSecurityManager.GetZoneMappings(uint dwZone, out IEnumString ppenumString, uint dwFlags)
        {
            ppenumString = null;
            return (int)WinInetErrors.INET_E_DEFAULT_ACTION;
        }

        #endregion
    }
}
