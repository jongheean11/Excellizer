using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Excellizer.Tab;

namespace Excellizer
{
    public partial class ThisAddIn
    {
        MainRibbon Ribbon;
        public bool chromeButtonChosen = false, IEButtonChosen = false, firefoxButtonChosen = false;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        protected override object RequestService(Guid serviceGuid)
        {
            if (serviceGuid == typeof(Office.IRibbonExtensibility).GUID)
            {
                if (Ribbon == null)
                    Ribbon = new MainRibbon();
                return Ribbon;
            }
            return base.RequestService(serviceGuid);
        }

        public Excel.Worksheet GetActiveWorksheet()
        {
            return ((Excel.Worksheet)Application.ActiveSheet);
        }
    }
}
