using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Excellizer;
using Excellizer.Model;
using System.Runtime.InteropServices;
using System.Net;
using mshtml;
using System.Diagnostics;
using Microsoft.Win32;

namespace Excellizer.Control
{
    public partial class BrowserForm : Form
    {
        /**
         * //Page Initializer
         * Scope Applier
         * Level Tracker
         * Submitter
         * -- 예제 --
         * HtmlDocument htmlDocument = webBrowser.Document;
            HtmlElement head = htmlDocument.GetElementsByTagName("head")[0];
            //HtmlDocument a = _webBrowser.Document.DomDocument as HtmlDocument;

            HtmlElement script = htmlDocument.CreateElement("script");
            //script.SetAttribute("text", "$(document).ready(function() {"
            //   + "alert('hello');"
            //    + "});");
            //script.SetAttribute("text", "function doHello() { alert('hello'); }");
            //head.AppendChild(script);
            //webBrowser.Document.InvokeScript("doHello");
         */
        private String string_ScopeApplier = "function ScopeApplier(){"
            + "$('')."
            + ""
            + "}";
        private String string_LevelTracker;
        private String string_Submitter;

        private Dictionary<HtmlElement, List<HtmlElement>> elementListDictionary;
        private Dictionary<HtmlElement, int> elementLevelDictionary;
        //private Dictionary<IHTMLElement, List<IHTMLElement>> elementListDictionary_MSHTML;
        private Dictionary<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>> elementDicDictionary_MSHTML;
        private Dictionary<IHTMLElement, int> elementLevelDictionary_MSHTML;

        private HtmlDocument prevHtmlDocument;
        private bool init = false;
        private bool afterBody = false;

        private List<List<HtmlElement>> detectedTable;

        private Dictionary<HtmlElement, Plexiglass> plexiglassDictionary;
        private Dictionary<IHTMLElement, Plexiglass> plexiglassDictionary_MSHTML;

        private Dictionary<Button, HtmlElement> buttonTargetDictionary;
        private Dictionary<Button, IHTMLElement> buttonTargetDictionary_MSHTML;
        private Point docLocation;

        AlertForm alert_Init;

        public BrowserForm()
        {
            InitializeComponent();
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.ObjectForScripting = true;
            elementListDictionary = new Dictionary<HtmlElement, List<HtmlElement>>();
            elementLevelDictionary = new Dictionary<HtmlElement, int>();
            //elementListDictionary_MSHTML = new Dictionary<IHTMLElement, List<IHTMLElement>>();
            elementDicDictionary_MSHTML = new Dictionary<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>>();
            elementLevelDictionary_MSHTML = new Dictionary<IHTMLElement, int>();
            buttonTargetDictionary = new Dictionary<Button, HtmlElement>();
            buttonTargetDictionary_MSHTML = new Dictionary<Button, IHTMLElement>();
            plexiglassDictionary  = new Dictionary<HtmlElement, Plexiglass>();
            plexiglassDictionary_MSHTML = new Dictionary<IHTMLElement, Plexiglass>();
            detectedTable = new List<List<HtmlElement>>();

            this.SizeChanged += BrowserForm_SizeChanged;
            this.MinimumSize = new Size(600, 600);
        }

        private void BrowserForm_Load(object sender, EventArgs e)
        {
            prevSize = this.Size;
            var appName = Process.GetCurrentProcess().ProcessName + ".exe";
            SetIE11KeyforWebBrowserControl(appName);
        }

        private void SetIE11KeyforWebBrowserControl(string appName)
        {
            RegistryKey Regkey = null;
            try
            {
                //For 64 bit Machine 
                if (Environment.Is64BitOperatingSystem)
                    Regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                else  //For 32 bit Machine 
                    Regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);

                //If the path is not correct or 
                //If user't have priviledges to access registry 
                if (Regkey == null)
                {
                    MessageBox.Show("Application Settings Failed - Address Not found");
                    return;
                }

                string FindAppkey = Convert.ToString(Regkey.GetValue(appName));

                //Check if key is already present 
                if (FindAppkey == "11001")
                {
                    //MessageBox.Show("Required Application Settings Present");
                    Regkey.Close();
                    return;
                }
                Regkey.SetValue(appName, unchecked((int)0x2AF9), RegistryValueKind.DWord);
                //If key is not present add the key , Kev value 8000-Decimal 
                if (string.IsNullOrEmpty(FindAppkey))
                    Regkey.SetValue(appName, unchecked((int)0x2AF9), RegistryValueKind.DWord);

                //check for the key after adding 
                FindAppkey = Convert.ToString(Regkey.GetValue(appName));

                if (FindAppkey != "11001")//(FindAppkey == "11001")
                    //MessageBox.Show("Application Settings Applied Successfully");
                //else
                    MessageBox.Show("Application Settings Failed, Ref: " + FindAppkey);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Application Settings Failed");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //Close the Registry 
                if (Regkey != null)
                    Regkey.Close();
            }
        }

        private Size prevSize;

        void BrowserForm_SizeChanged(object sender, System.EventArgs e)
        {
            toolStripTextBox_URL.Width = this.Size.Width - 441;

            if(prevSize.Width - this.Size.Width != 0)
            {
                if (buttonTargetDictionary.Count + buttonTargetDictionary_MSHTML.Count > 0)
                {
                    InitializeContents();
                }
            }
            prevSize = this.Size;
        }
        
        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //if (init)
            toolStripTextBox_URL.Text = webBrowser.Url.ToString();
            if (this.webBrowser.ReadyState == WebBrowserReadyState.Complete)
            {
                InitializeContents();
                HtmlDocument htmlDocument = webBrowser.Document;
                HtmlElement head = htmlDocument.GetElementsByTagName("head")[0];

                HtmlElement meta = htmlDocument.CreateElement("meta");
                meta.SetAttribute("text", "<meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />");
                head.AppendChild(meta);
            }
        }

        void InitializeView()
        {
            progressValue = 0;
            if (backgroundWorker_Init.IsBusy != true)
            {
                // create a new instance of the alert form
                alert_Init = new AlertForm();
                // event handler for the Cancel button in AlertForm
                ///alert_Init.Canceled += new EventHandler<EventArgs>(alertInitCancelButton_Click);
                alert_Init.Show();
                // Start the asynchronous operation.
                backgroundWorker_Init.RunWorkerAsync();
            }
            progressValue = 10;
            MakeStructure(webBrowser.Document.All);
            progressValue = 30;

            toolStripTextBox_URL.Text = webBrowser.Url.ToString();
            progressValue = 50;
            //  init = false;

            webBrowser.Document.Window.AttachEventHandler("onscroll", OnScrollEventHandler);
            progressValue = 60;
            docLocation = new Point(webBrowser.Document.GetElementsByTagName("HTML")[0].ScrollLeft, 
                webBrowser.Document.GetElementsByTagName("HTML")[0].ScrollTop);
            progressValue = 70;
            FormButtons();
            progressValue = 100;
        }

        void InitializeContents()
        {
            elementListDictionary.Clear();
            elementLevelDictionary.Clear();
            //elementListDictionary_MSHTML.Clear();
            elementDicDictionary_MSHTML.Clear();
            elementLevelDictionary_MSHTML.Clear();
            detectedTable.Clear();
            foreach (Button btn in buttonTargetDictionary.Keys)
            {
                this.Controls.Remove(plexiglassDictionary[buttonTargetDictionary[btn]]);
                plexiglassDictionary[buttonTargetDictionary[btn]].Dispose();
                this.Controls.Remove(btn);
                btn.Dispose();
            }
            foreach (Button btn in buttonTargetDictionary_MSHTML.Keys)
            {
                this.Controls.Remove(plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[btn]]);
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[btn]].Dispose();
                this.Controls.Remove(btn);
                btn.Dispose();
            }
            buttonTargetDictionary.Clear();
            buttonTargetDictionary_MSHTML.Clear();
            plexiglassDictionary.Clear();
            plexiglassDictionary_MSHTML.Clear();
            

            this.ResumeLayout();
        }

        private Timer timer1;
        public void InitTimer()
        {
            if (timer1 != null)
            {
                timer1.Stop();
                timer1.Dispose();
            }
            timer1 = new Timer();
            timer1.Tick += new EventHandler(timer1_Tick);
            timer1.Interval = 2000; // in miliseconds
            timer1.Start();
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            InitializeView();
        }

        #region HTML Data Linkify

        void MakeStructure(HtmlElementCollection htmleleCollection)
        {
            int iframeIndex = 0;
            foreach (HtmlElement htmlele in htmleleCollection)
            {
                if (afterBody)
                {
                    if (htmlele.TagName.Equals("IFRAME"))
                    {
                        StyleGenerator sg = new StyleGenerator();
                        sg.ParseStyleString(htmlele.Style == null ? "" : htmlele.Style);
                        if (!(sg.GetStyle("DISPLAY").Equals("none") || sg.GetStyle("VISIBILITY").Equals("hidden")
                            || sg.GetStyle("WIDTH").Equals("0") || sg.GetStyle("HEIGHT").Equals("0")
                            || sg.GetStyle("WIDTH").Equals("0px") || sg.GetStyle("HEIGHT").Equals("0px")
                            || sg.GetStyle("FILTER").Contains("opacity=0") || sg.GetStyle("-ms-filter").Contains("opacity=0")))
                        {
                            HTMLDocument htmldoc = webBrowser.Document.DomDocument as HTMLDocument;
                            IHTMLWindow2 frame = (IHTMLWindow2)htmldoc.frames.item(iframeIndex);

                            try
                            {
                                HTMLDocument doc2 = (HTMLDocument)frame.document;
                                MakeStructure_MSHTML(htmlele.DomElement as IHTMLElement, doc2.documentElement.all);
                            }
                            catch(System.UnauthorizedAccessException e)
                            {
                                e.ToString();
                            }
                        }
                        iframeIndex++;
                    }
                    else if (!((htmlele.TagName.Equals("SCRIPT")) || (htmlele.TagName.Equals("STYLE")) || (htmlele.TagName.Equals("!"))))
                    {
                        if (htmlele.Style != null)
                        {
                            String style = htmlele.Style.Replace(" ", String.Empty);
                            if (style.Contains("DISPLAY"))
                            {
                                int displayPos = style.IndexOf("DISPLAY:") + 8;
                                //widthPos = style.IndexOf("WIDTH:") + 6,
                                //heightPos = style.IndexOf("HEIGHT:") + 7;

                                if (!style.Substring(displayPos, 4).Equals("none"))
                                {
                                    LinkElements(htmlele);
                                }
                            }
                            else
                            {
                                LinkElements(htmlele);
                            }
                        }
                        else
                        {
                            LinkElements(htmlele);
                        }
                    }
                }
                else if (htmlele.TagName.Equals("BODY"))
                {
                    afterBody = true;

                    elementLevelDictionary.Add(htmlele, 0);

                    LinkElements(htmlele);
                }
                
            }
            afterBody = false;
        }

        void MakeStructure_MSHTML(IHTMLElement htmleleParent, IHTMLElementCollection htmleleCollection)
        {
            int iframeIndex = 0;
            bool _afterBody = false;
            foreach (IHTMLElement htmlele_MSHTML in htmleleCollection)
            {
                if (_afterBody)
                {
                    if (htmlele_MSHTML.tagName.Equals("IFRAME"))
                    {
                        StyleGenerator sg = new StyleGenerator();
                        sg.ParseStyleString(htmlele_MSHTML.style.toString());
                        if (!(sg.GetStyle("DISPLAY").Equals("none") || sg.GetStyle("VISIBILITY").Equals("hidden")
                            || sg.GetStyle("WIDTH").Equals("0") || sg.GetStyle("HEIGHT").Equals("0")
                            || sg.GetStyle("WIDTH").Equals("0px") || sg.GetStyle("HEIGHT").Equals("0px")
                            || sg.GetStyle("FILTER").Contains("opacity=0") || sg.GetStyle("-ms-filter").Contains("opacity=0")))
                        {
                            HTMLDocument htmldoc = htmlele_MSHTML.document as HTMLDocument;
                            IHTMLWindow2 frame = (IHTMLWindow2)htmldoc.frames.item(iframeIndex);

                            try
                            {
                                HTMLDocument doc2 = (HTMLDocument)frame.document;
                                MakeStructure_MSHTML(htmlele_MSHTML, doc2.documentElement.all);
                            }
                            catch (System.UnauthorizedAccessException e)
                            {
                                e.ToString();
                            }
                        }

                        iframeIndex++;
                    }
                    else if (!((htmlele_MSHTML.tagName.Equals("SCRIPT")) || (htmlele_MSHTML.tagName.Equals("STYLE")) || (htmlele_MSHTML.tagName.Equals("!"))))
                    {
                        if (htmlele_MSHTML.style != null)
                        {
                            String style = htmlele_MSHTML.style.toString().Replace(" ", String.Empty);
                            if (style.Contains("DISPLAY"))
                            {
                                int displayPos = style.IndexOf("DISPLAY:") + 8;
                                //widthPos = style.IndexOf("WIDTH:") + 6,
                                //heightPos = style.IndexOf("HEIGHT:") + 7;

                                if (!style.Substring(displayPos, 4).Equals("none"))
                                {
                                    LinkElements_MSHTML(htmlele_MSHTML, htmleleParent);
                                }
                            }
                            else
                            {
                                LinkElements_MSHTML(htmlele_MSHTML, htmleleParent);
                            }
                        }
                        else
                        {
                            LinkElements_MSHTML(htmlele_MSHTML, htmleleParent);
                        }
                    }
                }
                else if (htmlele_MSHTML.tagName.Equals("BODY"))
                {
                    _afterBody = true;

                    elementLevelDictionary_MSHTML.Add(htmlele_MSHTML, 0);

                    LinkElements_MSHTML(htmlele_MSHTML, htmleleParent);
                }
            }
            _afterBody = false;
        }

        private void LinkElements(HtmlElement htmlele)
        {
            List<HtmlElement> htmleleList = new List<HtmlElement>();
            foreach (HtmlElement _htmlele in htmlele.Children)
            {
                if (!((_htmlele.TagName.Equals("SCRIPT")) || (_htmlele.TagName.Equals("STYLE")) || (_htmlele.TagName.Equals("!"))))
                {
                    if (_htmlele.Style != null)
                    {
                        String style = _htmlele.Style.Replace(" ", String.Empty);
                        if (style.Contains("DISPLAY"))
                        {
                            int displayPos = style.IndexOf("DISPLAY:") + 8;
                            //widthPos = style.IndexOf("WIDTH:") + 6,
                            //heightPos = style.IndexOf("HEIGHT:") + 7;

                            if (!style.Substring(displayPos, 4).Equals("none"))
                            {
                                FormLevelAndElements(htmleleList, _htmlele);
                            }
                        }
                        else
                        {
                            FormLevelAndElements(htmleleList, _htmlele);
                        }
                    }
                    else
                    {
                        FormLevelAndElements(htmleleList, _htmlele);
                    }
                }
            }
            elementListDictionary.Add(htmlele, htmleleList);
        }

        private void LinkElements_MSHTML(IHTMLElement htmlele, IHTMLElement htmleleParent)
        {
            List<IHTMLElement> htmleleList = new List<IHTMLElement>();
            foreach (IHTMLElement _htmlele in htmlele.children)
            {
                if (!((_htmlele.tagName.Equals("SCRIPT")) || (_htmlele.tagName.Equals("STYLE")) || (_htmlele.tagName.Equals("!"))))
                {
                    if (_htmlele.style != null)
                    {
                        String style = _htmlele.style.toString().Replace(" ", String.Empty);
                        if (style.Contains("DISPLAY"))
                        {
                            int displayPos = style.IndexOf("DISPLAY:") + 8;

                            if (!style.Substring(displayPos, 4).Equals("none"))
                            {
                                FormLevelAndElements_MSHTML(htmleleList, _htmlele);
                            }
                        }
                        else
                        {
                            FormLevelAndElements_MSHTML(htmleleList, _htmlele);
                        }
                    }
                    else
                    {
                        FormLevelAndElements_MSHTML(htmleleList, _htmlele);
                    }
                }
            }
            //elementListDictionary_MSHTML.Add(htmlele, htmleleList);

            Dictionary<IHTMLElement, List<IHTMLElement>> elementListDictionary_MSHTML;
            if (!elementDicDictionary_MSHTML.ContainsKey(htmleleParent))
            {
                elementListDictionary_MSHTML = new Dictionary<IHTMLElement, List<IHTMLElement>>();
                elementListDictionary_MSHTML.Add(htmlele, htmleleList);
                elementDicDictionary_MSHTML.Add(htmleleParent, elementListDictionary_MSHTML);
            }
            else
            {
                elementListDictionary_MSHTML = elementDicDictionary_MSHTML[htmleleParent];
                elementListDictionary_MSHTML.Add(htmlele, htmleleList);
                elementDicDictionary_MSHTML[htmleleParent] = elementListDictionary_MSHTML;
            }
        }

        private void FormLevelAndElements(List<HtmlElement> htmleleList, HtmlElement _htmlele)
        {
            htmleleList.Add(_htmlele);

            HtmlElement current = _htmlele;
            int level=0;
            while(!current.TagName.Equals("BODY"))
            {
                level++;
                current = current.Parent;
            }
            if (elementLevelDictionary.ContainsKey(_htmlele))
                elementLevelDictionary.Remove(_htmlele);
            elementLevelDictionary.Add(_htmlele, level);
        }

        private void FormLevelAndElements_MSHTML(List<IHTMLElement> htmleleList, IHTMLElement _htmlele)
        {
            htmleleList.Add(_htmlele);

            IHTMLElement current = _htmlele;
            int level = 0;
            while (!current.tagName.Equals("BODY"))
            {
                level++;
                current = current.parentElement;
            }
            if (elementLevelDictionary_MSHTML.ContainsKey(_htmlele))
                elementLevelDictionary_MSHTML.Remove(_htmlele);
            elementLevelDictionary_MSHTML.Add(_htmlele, level);
        }

        #endregion

        #region Parse&Cancel Button

        private void parseButton_Click(object sender, EventArgs e)
        {
            foreach (KeyValuePair<Button, HtmlElement> kvPair in buttonTargetDictionary)
            {
                Button btn = kvPair.Key;
                HtmlElement htmlele = kvPair.Value;

                if (btn.BackColor == Color.LightSkyBlue)
                {
                    List<HtmlElement> htmleleList = elementListDictionary[htmlele];
                    int level = elementLevelDictionary[htmlele];

                    bool found = true;
                    int childCount = -1;

                    if (htmlele.TagName.Equals("TABLE"))
                    {
                        foreach (HtmlElement _htmlele in htmlele.Children)
                        {
                            if (!(_htmlele.TagName.Equals("COLGROUP") || _htmlele.TagName.Equals("ROWGROUP")))
                            {
                                if (_htmlele.Style != null)
                                {
                                    String style = _htmlele.Style.Replace(" ", String.Empty);
                                    if (style.Contains("DISPLAY"))
                                    {
                                        int displayPos = style.IndexOf("DISPLAY:") + 8;
                                        //widthPos = style.IndexOf("WIDTH:") + 6,
                                        //heightPos = style.IndexOf("HEIGHT:") + 7;

                                        if (!style.Substring(displayPos, 4).Equals("none"))
                                        {
                                            FormCells(_htmlele);
                                        }
                                    }
                                    else
                                    {
                                        FormCells(_htmlele);
                                    }
                                }
                                else
                                {
                                    FormCells(_htmlele);
                                }
                            }
                        }
                    }
                    else if (htmlele.TagName.Equals("UL") || htmlele.TagName.Equals("DL"))
                    {
                        foreach (HtmlElement _htmlele in htmlele.Children)
                        {
                            if (_htmlele.Style != null)
                            {
                                String style = _htmlele.Style.Replace(" ", String.Empty);
                                if (style.Contains("DISPLAY"))
                                {
                                    int displayPos = style.IndexOf("DISPLAY:") + 8;
                                    //widthPos = style.IndexOf("WIDTH:") + 6,
                                    //heightPos = style.IndexOf("HEIGHT:") + 7;

                                    if (!style.Substring(displayPos, 4).Equals("none"))
                                    {
                                        FormCells(_htmlele);
                                    }
                                }
                                else
                                {
                                    FormCells(_htmlele);
                                }
                            }
                            else
                            {
                                FormCells(_htmlele);
                            }
                        }
                    }
                    /*
                    else
                    {
                        foreach (HtmlElement _htmlele in htmleleList)
                        {
                            if (elementLevelDictionary.ContainsKey(_htmlele))
                            {
                                if (childCount == -1)
                                    childCount = elementListDictionary[_htmlele].Count;
                                else
                                {
                                    if (childCount != elementListDictionary[_htmlele].Count)
                                    {
                                        found = false;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                found = false;
                                break;
                            }
                        }
                    }
                    */
                    if (!found)
                    {
                        foreach (HtmlElement _htmlele in htmleleList)
                        {
                            FormCells(_htmlele);
                        }
                    }
                }
            }

            foreach (KeyValuePair<Button, IHTMLElement> kvPair in buttonTargetDictionary_MSHTML)
            {
                Button btn = kvPair.Key;
                IHTMLElement htmlele = kvPair.Value;

                if (btn.BackColor == Color.LightSkyBlue)
                {
                    //List<IHTMLElement> htmleleList = elementDicDictionary_MSHTML[htmlele];
                    int level = elementLevelDictionary_MSHTML[htmlele];

                    bool found = true;
                    int childCount = -1;

                    if (htmlele.tagName.Equals("TABLE"))
                    {
                        foreach (HtmlElement _htmlele in htmlele.children)
                        {
                            if (!(_htmlele.TagName.Equals("COLGROUP") || _htmlele.TagName.Equals("ROWGROUP")))
                            {
                                if (_htmlele.Style != null)
                                {
                                    String style = _htmlele.Style.Replace(" ", String.Empty);
                                    if (style.Contains("DISPLAY"))
                                    {
                                        int displayPos = style.IndexOf("DISPLAY:") + 8;
                                        //widthPos = style.IndexOf("WIDTH:") + 6,
                                        //heightPos = style.IndexOf("HEIGHT:") + 7;

                                        if (!style.Substring(displayPos, 4).Equals("none"))
                                        {
                                            FormCells(_htmlele);
                                        }
                                    }
                                    else
                                    {
                                        FormCells(_htmlele);
                                    }
                                }
                                else
                                {
                                    FormCells(_htmlele);
                                }
                            }
                        }
                    }
                    else if (htmlele.tagName.Equals("UL") || htmlele.tagName.Equals("DL"))
                    {
                        foreach (HtmlElement _htmlele in htmlele.children)
                        {
                            if (_htmlele.Style != null)
                            {
                                String style = _htmlele.Style.Replace(" ", String.Empty);
                                if (style.Contains("DISPLAY"))
                                {
                                    int displayPos = style.IndexOf("DISPLAY:") + 8;
                                    //widthPos = style.IndexOf("WIDTH:") + 6,
                                    //heightPos = style.IndexOf("HEIGHT:") + 7;

                                    if (!style.Substring(displayPos, 4).Equals("none"))
                                    {
                                        FormCells(_htmlele);
                                    }
                                }
                                else
                                {
                                    FormCells(_htmlele);
                                }
                            }
                            else
                            {
                                FormCells(_htmlele);
                            }
                        }
                    }
                    /*
                    else
                    {
                        foreach (HtmlElement _htmlele in htmleleList)
                        {
                            if (elementLevelDictionary.ContainsKey(_htmlele))
                            {
                                if (childCount == -1)
                                    childCount = elementListDictionary[_htmlele].Count;
                                else
                                {
                                    if (childCount != elementListDictionary[_htmlele].Count)
                                    {
                                        found = false;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                found = false;
                                break;
                            }
                        }
                    }
                    */
                    if (!found)
                    {
                        //foreach (HtmlElement _htmlele in htmleleList)
                        //{
                            //FormCells(_htmlele);
                        //}
                    }
                }
            }

            ////InsertDatas();
        }

        int maxColumnCount_Table, maxRowCount_Table;

        private void FormCells(HtmlElement _htmlele)
        {
            int count, maxCount = 0;
            maxRowCount_Table = 0;
            foreach (HtmlElement row in _htmlele.Children)
            {
                count = 0;
                List<HtmlElement> htmleleList = new List<HtmlElement>();
                foreach (HtmlElement cell in row.Children)
                {
                    htmleleList.Add(cell);
                    count++;
                }
                if (maxCount < count)
                    maxCount = count;
                if (htmleleList.Count != 0)
                {
                    detectedTable.Add(htmleleList);
                }
                maxRowCount_Table++;
            }
            maxColumnCount_Table = maxCount;

            var addins = Globals.ThisAddIn;
            Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)addins.Application.Worksheets.Add();

            InsertDatas();
            detectedTable.Clear();
        }

        private void InsertDatas()
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int endRow = detectedTable.Count;
            int endCol = GetEndColumn();
            int idxRow = 1, idxCol = 1, selectedX = activeCell.Row - 1, selectedY = activeCell.Column - 1;
            int rowspan = 1, colspan = 1, rowspan_check = 0, colspan_check = 0;

            StyleGenerator sg = new StyleGenerator();

            Dictionary<int, Dictionary<int, int>> checkMatrix = new Dictionary<int, Dictionary<int, int>>();
            for (int i = 0; i < maxRowCount_Table; i++)
            {
                Dictionary<int,int> tempDic = new Dictionary<int,int>();
                for(int j=0; j < maxColumnCount_Table; j++)
                {
                    tempDic.Add(j, 0);
                }
                checkMatrix.Add(i, tempDic);
            }

            foreach (List<HtmlElement> htmleleList in detectedTable)
            {
                idxCol = 1;
                foreach (HtmlElement htmlele in htmleleList)
                {
                    rowspan = 1;
                    colspan = 1;

                    bool _skip = false;
                    for (; idxCol <= maxColumnCount_Table;)
                    {
                        if (checkMatrix[idxRow - 1][idxCol - 1] == 1)
                        {
                            if (idxCol == maxColumnCount_Table)
                            {
                                _skip = true;
                                break;
                            }
                            else
                                idxCol++;
                        }
                        else
                            break;
                    }
                    if (_skip)
                        break;

                    sg.ParseStyleString(htmlele.Style == null ? "" : htmlele.Style);

                    if ((!sg.GetStyle("rowspan").Equals("")) && (!sg.GetStyle("colspan").Equals("")))
                    {
                        if ((!sg.GetStyle("rowspan").Equals("1")) && (!sg.GetStyle("colspan").Equals("1")))
                        {
                            rowspan = int.Parse(sg.GetStyle("rowspan"));
                            colspan = int.Parse(sg.GetStyle("colspan"));
                            activeSheet.Range[activeSheet.Cells[idxRow + selectedX, idxCol + selectedY], 
                                activeSheet.Cells[idxRow + rowspan - 1 + selectedX, idxCol + colspan - 1 + selectedY]].Merge();
                        }
                        else if (!sg.GetStyle("rowspan").Equals("1"))
                        {
                            rowspan = int.Parse(sg.GetStyle("rowspan"));
                            activeSheet.Range[activeSheet.Cells[idxRow + selectedX, idxCol + selectedY], 
                                activeSheet.Cells[idxRow + rowspan - 1 + selectedX, idxCol + selectedY]].Merge();
                        }
                        else if (!sg.GetStyle("colspan").Equals("1"))
                        {
                            colspan = int.Parse(sg.GetStyle("colspan"));
                            activeSheet.Range[activeSheet.Cells[idxRow + selectedX, idxCol + selectedY], 
                                activeSheet.Cells[idxRow + selectedX, idxCol + colspan - 1 + selectedY]].Merge();
                        }
                    }
                    else if (!sg.GetStyle("rowspan").Equals(""))
                    {
                        if (!sg.GetStyle("rowspan").Equals("1"))
                        {
                            rowspan = int.Parse(sg.GetStyle("rowspan"));
                            activeSheet.Range[activeSheet.Cells[idxRow + selectedX, idxCol + selectedY], 
                                activeSheet.Cells[idxRow + rowspan - 1 + selectedX, idxCol + selectedY]].Merge();
                        }
                    }
                    else if (!sg.GetStyle("colspan").Equals(""))
                    {
                        if (!sg.GetStyle("colspan").Equals("1"))
                        {
                            colspan = int.Parse(sg.GetStyle("colspan"));
                            activeSheet.Range[activeSheet.Cells[idxRow + selectedX, idxCol + selectedY],
                                activeSheet.Cells[idxRow + selectedX, idxCol + colspan - 1 + selectedY]].Merge();
                            //factorList.Add(new TableSeperatingFactor(idxRow, idxCol, 0, colspan));
                        }
                    }

                    for (int i = idxRow - 1; i < (idxRow - 1 + rowspan); i++)
                    {
                        for (int j = idxCol - 1; j< (idxCol - 1 + colspan); j++)
                        {
                            checkMatrix[i][j] = 1;
                        }
                    }
                    //((Excel.Range)activeSheet.Cells[idxRow, idxCol]).Font = htmlele.GetAttribute("")
                    ((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]).Value2 = htmlele.InnerHtml;
                    
                    idxCol = idxCol + colspan;
                }
               
                idxRow++;
            }
        }

        private int GetEndColumn()
        {
            int maxCol = 0;
            foreach (List<HtmlElement> htmleleList in detectedTable)
            {
                if (maxCol < htmleleList.Count)
                    maxCol = htmleleList.Count;
            }
            return maxCol;
        }

        #endregion

        #region 상단 툴팁 컨트롤 이벤트 함수

        private void toolStripButton_Move_Click(object sender, System.EventArgs e)
        {
            init = true;
            /*
            int size = 0;
            StringBuilder lpszCookieData = new StringBuilder(size);
            InternetGetCookie("http://beanbox.azurewebsites.net", null, ref lpszCookieData, ref size);

            // get cookie
            string cookie = lpszCookieData.ToString();

            InternetSetCookie("http://beanbox.azurewebsites.net", null, cookie);
            */

            //int _index = urlTextBox.Text.IndexOf("://");
            //string _path = urlTextBox.Text.Substring(_index+3+urlTextBox.Text.Substring(_index+3).IndexOf("/")+1);
            //Cookie cookie = new Cookie(urlTextBox.Text, urlTextBox.Text, _path, "/");
            //InternetSetCookie(urlTextBox.Text, null, cookie.ToString() + "; expires = Sun, 01-Jan-2013 00:00:00 GMT");
            if (!toolStripTextBox_URL.Text.StartsWith("http://") && !toolStripTextBox_URL.Text.StartsWith("https://"))
                toolStripTextBox_URL.Text = "http://" + toolStripTextBox_URL.Text;
            //InitializeContents();
            webBrowser.Navigate(toolStripTextBox_URL.Text);
        }

        private void toolStripTextBox_URL_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                if (toolStripTextBox_URL.Text.Equals(""))
                    toolStripTextBox_URL.Text = "http://localhost:8080/Performance";
                init = true;
                if (!toolStripTextBox_URL.Text.StartsWith("http://") && !toolStripTextBox_URL.Text.StartsWith("https://"))
                    toolStripTextBox_URL.Text = "http://" + toolStripTextBox_URL.Text;
                //InitializeContents();
                webBrowser.Navigate(toolStripTextBox_URL.Text);
            }
        }

        private void toolStripButton_Detect_Click(object sender, EventArgs e)
        {
            if (buttonTargetDictionary.Count + buttonTargetDictionary.Count > 0)
                InitializeContents();
            InitializeView();
        }

        private void toolStripButton_Back_Click(object sender, EventArgs e)
        {
            webBrowser.GoBack();
        }

        private void toolStripButton_Forward_Click(object sender, EventArgs e)
        {
            webBrowser.GoForward();
        }

        private void toolStripButton_Home_Click(object sender, EventArgs e)
        {
            webBrowser.GoHome();
        }

        private void toolStripButton_Refresh_Click(object sender, EventArgs e)
        {
            webBrowser.Refresh();
        }

        private void toolStripButton_Stop_Click(object sender, EventArgs e)
        {
            webBrowser.Stop();
        }
        #endregion

        public void OnScrollEventHandler(object sender, EventArgs e)
        {
            int scrollTop = webBrowser.Document.GetElementsByTagName("HTML")[0].ScrollTop,
                scrollLeft = webBrowser.Document.GetElementsByTagName("HTML")[0].ScrollLeft;
            int xDiff = docLocation.X - scrollLeft, yDiff = docLocation.Y - scrollTop;
            
            docLocation.X = scrollLeft;
            docLocation.Y = scrollTop;

            foreach (Button btn in buttonTargetDictionary.Keys)
            {
                btn.Location = new Point(btn.Location.X + xDiff, btn.Location.Y + yDiff);
                if ((btn.Location.X >= 0) && (btn.Location.X <= webBrowser.Width - 20)
                    && (btn.Location.Y >= 33) && (btn.Location.Y <= webBrowser.Height + 13))
                {
                    btn.Visible = true;
                }
                else
                {
                    btn.Visible = false;
                }
            }

            foreach (Button btn in buttonTargetDictionary_MSHTML.Keys)
            {
                btn.Location = new Point(btn.Location.X + xDiff, btn.Location.Y + yDiff);
                if ((btn.Location.X >= 0) && (btn.Location.X <= webBrowser.Width - 20)
                    && (btn.Location.Y >= 33) && (btn.Location.Y <= webBrowser.Height + 13))
                {
                    btn.Visible = true;
                }
                else
                {
                    btn.Visible = false;
                }
            }

            Point parentPtr = this.PointToScreen(Point.Empty);
            foreach (KeyValuePair<HtmlElement, Plexiglass> kvPair in plexiglassDictionary)
            {
                Plexiglass plexiglass = kvPair.Value;
                HtmlElement htmlele = kvPair.Key;

                int x = plexiglass.Location.X, y = plexiglass.Location.Y, 
                    width = plexiglass.Width, height = plexiglass.Height;
                if (xDiff > 0)
                {
                    if ((plexiglass._offsetLeft > docLocation.X) && (plexiglass._offsetLeft < webBrowser.Width + docLocation.X))
                    {
                        x = plexiglass._offsetLeft - (docLocation.X);
                        plexiglass.Location = new Point(x + parentPtr.X, y);
                        plexiglass.ClientSize = new Size(plexiglass._width, plexiglass.ClientSize.Height);
                        if (plexiglass._offsetLeft - docLocation.X > webBrowser.Width)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary.FirstOrDefault(kv => kv.Value == htmlele).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                    else if (plexiglass._offsetLeft > webBrowser.Width + docLocation.X)
                    {
                        plexiglass.Visible = false;
                    }
                    else
                    {
                        x = 0;
                        plexiglass.Location = new Point(x + parentPtr.X, y);
                        int applyWidth = plexiglass._width - (docLocation.X - plexiglass._offsetLeft);
                        plexiglass.ClientSize = new Size(applyWidth, plexiglass.ClientSize.Height); // plexiglass.ClientSize.Height - (docLocation.Y - (plexiglass._y - 33)));
                        if (applyWidth <= 0)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary.FirstOrDefault(kv => kv.Value == htmlele).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                }
                else
                {
                    if ((plexiglass._offsetTop > docLocation.Y) && (plexiglass._offsetTop < webBrowser.Height + docLocation.Y))
                    {
                        y = plexiglass._offsetTop + 33 - (docLocation.Y);
                        plexiglass.Location = new Point(x, y + parentPtr.Y);
                        plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, plexiglass._height);
                        if (plexiglass._offsetTop - docLocation.Y > webBrowser.Height)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary.FirstOrDefault(kv => kv.Value == htmlele).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                    else if (plexiglass._offsetTop > webBrowser.Height + docLocation.Y)
                    {
                        plexiglass.Visible = false;
                    }
                    else
                    {
                        y = 35;
                        plexiglass.Location = new Point(x, y + parentPtr.Y);
                        int applyHeight = plexiglass._height - (docLocation.Y - plexiglass._offsetTop) + 10;
                        plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, applyHeight); // plexiglass.ClientSize.Height - (docLocation.Y - (plexiglass._y - 33)));
                        if (applyHeight <= 0)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary.FirstOrDefault(kv => kv.Value == htmlele).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                }
            }
            foreach (KeyValuePair<IHTMLElement, Plexiglass> kvPair in plexiglassDictionary_MSHTML)
            {
                Plexiglass plexiglass = kvPair.Value;
                IHTMLElement htmlele_MSHTML = kvPair.Key;

                int x = plexiglass.Location.X, y = plexiglass.Location.Y,
                    width = plexiglass.Width, height = plexiglass.Height;
                if (xDiff > 0)
                {
                    if ((plexiglass._offsetLeft > docLocation.X) && (plexiglass._offsetLeft < webBrowser.Width + docLocation.X))
                    {
                        x = plexiglass._offsetLeft - (docLocation.X);
                        plexiglass.Location = new Point(x + parentPtr.X, y);
                        plexiglass.ClientSize = new Size(plexiglass._width, plexiglass.ClientSize.Height);
                        if (plexiglass._offsetLeft - docLocation.X > webBrowser.Width)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary_MSHTML.FirstOrDefault(kv => kv.Value == htmlele_MSHTML).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                    else if (plexiglass._offsetLeft > webBrowser.Width + docLocation.X)
                    {
                        plexiglass.Visible = false;
                    }
                    else
                    {
                        x = 0;
                        plexiglass.Location = new Point(x + parentPtr.X, y);
                        int applyWidth = plexiglass._width - (docLocation.X - plexiglass._offsetLeft);
                        plexiglass.ClientSize = new Size(applyWidth, plexiglass.ClientSize.Height); // plexiglass.ClientSize.Height - (docLocation.Y - (plexiglass._y - 33)));
                        if (applyWidth <= 0)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary_MSHTML.FirstOrDefault(kv => kv.Value == htmlele_MSHTML).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                }
                else
                {
                    if ((plexiglass._offsetTop > docLocation.Y) && (plexiglass._offsetTop < webBrowser.Height + docLocation.Y))
                    {
                        if ((plexiglass.BackColor == Color.LightSkyBlue) && plexiglass.Visible)
                            y = 0;
                        //y = plexiglass.Location.Y - parentPtr.Y + yDiff;
                        y = plexiglass._offsetTop + 33 - docLocation.Y;
                        plexiglass.Location = new Point(x, y + parentPtr.Y);
                        plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, plexiglass._height);
                        if (plexiglass._offsetTop - docLocation.Y > webBrowser.Height)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary_MSHTML.FirstOrDefault(kv => kv.Value == htmlele_MSHTML).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                    else if (plexiglass._offsetTop > webBrowser.Height + docLocation.Y)
                    {
                        plexiglass.Visible = false;
                    }
                    else
                    {
                        y = 33;
                        plexiglass.Location = new Point(x, y + parentPtr.Y);
                        int applyHeight = plexiglass._height - (docLocation.Y - plexiglass._offsetTop);
                        plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, applyHeight); // plexiglass.ClientSize.Height - (docLocation.Y - (plexiglass._y - 33)));
                        if (applyHeight <= 0)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary_MSHTML.FirstOrDefault(kv => kv.Value == htmlele_MSHTML).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                }
            }

            this.ResumeLayout();
        }

        #region 영역 및 선택 버튼 생성 함수

        private void FormButtons()
        {
            foreach (HtmlElement htmlele in elementListDictionary.Keys)
            {
                if (htmlele.TagName.Equals("TABLE"))
                {
                    CreateButton(htmlele);
                }
            }
            foreach (KeyValuePair<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>> kvPair in elementDicDictionary_MSHTML)
            {
                IHTMLElement htmleleParent = kvPair.Key;
                Dictionary<IHTMLElement, List<IHTMLElement>> elementListDictionary_MSHTML = kvPair.Value;
                foreach (IHTMLElement htmlele_MSHTML in elementListDictionary_MSHTML.Keys)
                {
                    if (htmlele_MSHTML.tagName.Equals("TABLE"))
                    {
                        CreateButton_MSHTML(htmlele_MSHTML, htmleleParent);
                    }
                }
            }

            this.ResumeLayout();
        }

        private void SetElementXY(HtmlElement htmlele, out int x, out int y)
        {
            // Calculate the offset of the element, all the way up through the parent nodes
            var parent = htmlele.OffsetParent;
            int xoff = htmlele.OffsetRectangle.X;
            int yoff = htmlele.OffsetRectangle.Y;

            while (parent != null)
            {
                xoff += parent.OffsetRectangle.X;
                yoff += parent.OffsetRectangle.Y;
                parent = parent.OffsetParent;
            }

            x = xoff; y = yoff;
            // Get the scrollbar offsets
            ////int scrollBarYPosition = docLocation.X;
            ////int scrollBarXPosition = docLocation.Y;

            // Calculate the visible page space
            ////Rectangle visibleWindow = new Rectangle(scrollBarXPosition, scrollBarYPosition, webBrowser.Width, webBrowser.Height);

            // Calculate the visible area of the element
            //// Rectangle elementWindow = new Rectangle(xoff, yoff, htmlele.ClientRectangle.Width, htmlele.ClientRectangle.Height);
        }

        private void SetElementXY_MSHTML(IHTMLElement htmlele_MSHTML, out int x, out int y)
        {
            // Calculate the offset of the element, all the way up through the parent nodes
            var parent = htmlele_MSHTML.offsetParent;
            int xoff = htmlele_MSHTML.offsetLeft;
            int yoff = htmlele_MSHTML.offsetTop;

            while (parent != null)
            {
                xoff += parent.offsetLeft;
                yoff += parent.offsetTop;
                parent = parent.offsetParent;
            }

            x = xoff; y = yoff;
        }

        private void CreateButton(HtmlElement htmlele)
        {
            int x, y;
            SetElementXY(htmlele, out x, out y);
            
            Button newButton = new Button();
            Point point = new Point(x < 20 ? 0 - docLocation.X : x - 20 - docLocation.X, y < 20 ? 0 - docLocation.Y : y - 20 - docLocation.Y);
            newButton.Location = point;
            newButton.Width = 15;
            newButton.Height = 15;
            newButton.Text = "V";
            newButton.BackColor = Color.LightGray;

            if((docLocation.Y <= y) && (y + 20 <= webBrowser.Height))
            {
                newButton.Visible = true;
            }
            else
            {
                newButton.Visible = false;
            }

            newButton.Click += newButton_Click;
            newButton.MouseEnter += newButton_MouseEnter;
            newButton.MouseLeave += newButton_MouseLeave;
            this.Controls.Add(newButton);
            newButton.Parent = webBrowser;

            buttonTargetDictionary.Add(newButton, htmlele);

            Plexiglass plexiglass = new Plexiglass(this, x - docLocation.X, y + 33 - docLocation.Y, htmlele.OffsetRectangle.Width, htmlele.OffsetRectangle.Height);
            plexiglass._offsetLeft = x;
            plexiglass._offsetTop = y;
            plexiglassDictionary.Add(htmlele, plexiglass);
        }

        private void CreateButton_MSHTML(IHTMLElement htmlele_MSHTML, IHTMLElement htmleleParent)
        {
            int x, y;
            SetElementXY_MSHTML(htmlele_MSHTML, out x, out y);
            int parentX, parentY;
            SetElementXY_MSHTML(htmleleParent, out parentX, out parentY);

            Button newButton = new Button();
            Point point = new Point(x + parentX < 20 ? 0 - docLocation.X : x + parentX - 20 - docLocation.X, y + parentY < 20 ? 0 - docLocation.Y : y + parentY - 20 - docLocation.Y);
            newButton.Location = point;
            newButton.Width = 15;
            newButton.Height = 15;
            newButton.Text = "V";
            newButton.BackColor = Color.LightGray;

            if ((docLocation.Y <= y + parentY) && (y + parentY + 20 <= webBrowser.Height))
            {
                newButton.Visible = true;
            }
            else
            {
                newButton.Visible = false;
            }

            newButton.Click += newButton_Click_MSHTML;
            newButton.MouseEnter += newButton_MouseEnter_MSHTML;
            newButton.MouseLeave += newButton_MouseLeave_MSHTML;
            this.Controls.Add(newButton);
            newButton.Parent = webBrowser;

            buttonTargetDictionary_MSHTML.Add(newButton, htmlele_MSHTML);

            Plexiglass plexiglass = new Plexiglass(this,
                x + parentX - docLocation.X, y + parentY + 33 - docLocation.Y, 
                htmlele_MSHTML.offsetWidth, htmlele_MSHTML.offsetHeight);
            plexiglass._offsetLeft = x + parentX;
            plexiglass._offsetTop = y + parentY;
            plexiglassDictionary_MSHTML.Add(htmlele_MSHTML, plexiglass);
        }

        #endregion

        #region 영역 선택 버튼 이벤트 함수

        void newButton_MouseLeave(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                plexiglassDictionary[buttonTargetDictionary[newButton]].BackColor = Color.LightSkyBlue;
                plexiglassDictionary[buttonTargetDictionary[newButton]].Visible = false;
            }
            else
            {
                plexiglassDictionary[buttonTargetDictionary[newButton]].BackColor = Color.LightSkyBlue;
            }
        }

        void newButton_MouseLeave_MSHTML(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].BackColor = Color.LightSkyBlue;
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].Visible = false;
            }
            else
            {
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].BackColor = Color.LightSkyBlue;
            }
        }

        void newButton_MouseEnter(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                plexiglassDictionary[buttonTargetDictionary[newButton]].BackColor = Color.LightBlue;
                plexiglassDictionary[buttonTargetDictionary[newButton]].Visible = true;
            }
            else
            {
                plexiglassDictionary[buttonTargetDictionary[newButton]].BackColor = Color.LightBlue;
            }
            this.ResumeLayout();
        }

        void newButton_MouseEnter_MSHTML(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].BackColor = Color.LightBlue;
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].Visible = true;
            }
            else
            {
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].BackColor = Color.LightBlue;
            }
            this.ResumeLayout();
        }

        void newButton_Click(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if(newButton.BackColor.Equals(Color.LightGray))
            {
                newButton.BackColor = Color.Aqua;
                plexiglassDictionary[buttonTargetDictionary[newButton]].Visible = true;
            }
            else
            {
                newButton.BackColor = Color.LightGray;
                plexiglassDictionary[buttonTargetDictionary[newButton]].Visible = false;
            }
        }

        void newButton_Click_MSHTML(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                newButton.BackColor = Color.Aqua;
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].Visible = true;
            }
            else
            {
                newButton.BackColor = Color.LightGray;
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].Visible = false;
            }
        }

#endregion

        #region waiting progress bar

        int progressValue = 0;

        private void backgroundWorker_Init_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            int _value = 0;
            while(_value != 100)
            {
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    if (_value != progressValue)
                    {
                        _value = progressValue;
                        worker.ReportProgress(_value);
                        System.Threading.Thread.Sleep(500);
                    }
                }
            }
        }

        // This event handler updates the progress.
        private void backgroundWorker_Init_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Show the progress in main form (GUI)
            /// labelResult.Text = (e.ProgressPercentage.ToString() + "%");
            // Pass the progress to AlertForm label and progressbar
            alert_Init.Message = "In progress, please wait... " + e.ProgressPercentage.ToString() + "%";
            alert_Init.ProgressValue = e.ProgressPercentage;
        }

        // This event handler deals with the results of the background operation.
        private void backgroundWorker_Init_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            /*if (e.Cancelled == true)
            {
                labelResult.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                labelResult.Text = "Error: " + e.Error.Message;
            }
            else
            {
                labelResult.Text = "Done!";
            }
            *///
            // Close the AlertForm
            alert_Init.Close();
        }

        public event EventHandler<EventArgs> Canceled;

        private void alertInitCancelButton_Click(object sender, EventArgs e)
        {
            // Create a copy of the event to work with
            EventHandler<EventArgs> ea = Canceled;
            /* If there are no subscribers, eh will be null so we need to check
             * to avoid a NullReferenceException. */
            if (ea != null)
                ea(this, e);
        }

        #endregion

        #region etc(not using)

        [DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetGetCookieEx(
            string url,
            string cookieName,
            StringBuilder cookieData,
            ref int size,
            Int32 dwFlags,
            IntPtr lpReserved);

        private const Int32 InternetCookieHttponly = 0x2000;

        /// <summary>
        /// Gets the URI cookie container.
        /// </summary>
        /// <param name="uri">The URI.</param>
        /// <returns></returns>
        public static CookieContainer GetUriCookieContainer(Uri uri)
        {
            CookieContainer cookies = null;
            // Determine the size of the cookie
            int datasize = 8192 * 16;
            StringBuilder cookieData = new StringBuilder(datasize);
            if (!InternetGetCookieEx(uri.ToString(), null, cookieData, ref datasize, InternetCookieHttponly, IntPtr.Zero))
            {
                if (datasize < 0)
                    return null;
                // Allocate stringbuilder large enough to hold the cookie
                cookieData = new StringBuilder(datasize);
                if (!InternetGetCookieEx(uri.ToString(), null, cookieData, ref datasize, InternetCookieHttponly, IntPtr.Zero))
                    return null;
            }
            if (cookieData.Length > 0)
            {
                cookies = new CookieContainer();
                cookies.SetCookies(uri, cookieData.ToString().Replace(';', ','));
            }
            return cookies;
        }

        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetSetCookie(string lpszUrlName, string lpszCookieName, string lpszCookieData);

        [DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetGetCookie(string lpszUrl, string lpszCookieName, ref StringBuilder lpszCookieData, ref int lpdwSize);

        private void button1_Click(object sender, EventArgs e)
        {
            //foreach (HtmlElement htmlele in elementListDictionary.Keys)
            /*foreach (HtmlElement htmlele in webBrowser.Document.All)
            {
                //int width = htmlele.ClientRectangle.Width;
                //int height = htmlele.ClientRectangle.Height;
                int width = htmlele.OffsetRectangle.Width;
                int height = htmlele.OffsetRectangle.Height;
                if ((width != 0) || (height != 0))
                {
                    continue;
                }
            }*/

            InitializeView();
            /*
            StyleGenerator sg = new StyleGenerator();

            foreach (HtmlElement htmlele in webBrowser.Document.All)
            {
                if (htmlele.TagName.Equals("TABLE"))
                {
                    Plexiglass plexiglass = new Plexiglass(this, htmlele.OffsetRectangle.X + 10 - docLocation.X, htmlele.OffsetRectangle.Y + 48 - docLocation.Y, htmlele.OffsetRectangle.Width, htmlele.OffsetRectangle.Height);
                    
                    plexiglass.Visible = false;

                    plexiglassDictionary.Add(htmlele, plexiglass);
                }
            }

            foreach (KeyValuePair<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>> kvPair in elementDicDictionary_MSHTML)
            {
                IHTMLElement htmleleParent = kvPair.Key;
                Dictionary<IHTMLElement, List<IHTMLElement>> elementListDictionary_MSHTML = kvPair.Value;
                foreach(IHTMLElement htmlele_MSHTML in elementListDictionary_MSHTML.Keys)
                {
                    if (htmlele_MSHTML.tagName.Equals("TABLE"))
                    {
                        Plexiglass plexiglass = new Plexiglass(this, 
                            htmleleParent.offsetLeft + 10 - docLocation.X + htmlele_MSHTML.offsetLeft,
                            htmleleParent.offsetTop + 48 - docLocation.Y + htmlele_MSHTML.offsetTop, 
                            htmlele_MSHTML.offsetWidth, htmlele_MSHTML.offsetHeight);

                        plexiglass.Visible = false;

                        plexiglassDictionary_MSHTML.Add(htmlele_MSHTML, plexiglass);
                    }
                }
            }
             * */
        }

        #endregion

    }
}