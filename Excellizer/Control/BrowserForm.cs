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
using System.Text.RegularExpressions;

namespace Excellizer.Control
{
    public partial class BrowserForm : Form
    {
        #region Attributes and Special Controls

        private Dictionary<HtmlElement, List<HtmlElement>> elementListDictionary;
        //private Dictionary<HtmlElement, int> elementLevelDictionary;
        private Dictionary<int, List<HtmlElement>> elementLevelDictionary;
        //private Dictionary<IHTMLElement, List<IHTMLElement>> elementListDictionary_MSHTML;
        private Dictionary<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>> elementDicDictionary_MSHTML;
        private Dictionary<IHTMLElement, int> parentFrameIndexDictionary;
        private Dictionary<IHTMLElement, int> elementLevelDictionary_MSHTML;

        private bool init = false;
        private bool afterBody = false;

        private List<List<HtmlElement>> detectedTable;
        private List<List<IHTMLElement>> detectedTable_MSHTML;

        private Dictionary<HtmlElement, Plexiglass> plexiglassDictionary;
        private Dictionary<IHTMLElement, Plexiglass> plexiglassDictionary_MSHTML;
        private Dictionary<List<List<HtmlElement>>, Plexiglass> plexiglassDictionary_Level;

        private int regionSelected;
        private Dictionary<Button, HtmlElement> buttonTargetDictionary;
        private Dictionary<Button, IHTMLElement> buttonTargetDictionary_MSHTML;
        private Dictionary<Button, List<List<HtmlElement>>> buttonTargetDictionary_Level;
        private Point docLocation;

        private AlertForm alert_Init;

        private MultiPageSettingForm mulForm;
        
        #endregion

        #region BrowserForm 이벤트 메서드

        public BrowserForm()
        {
            InitializeComponent();
            webBrowser.ScriptErrorsSuppressed = true;
            webBrowser.ObjectForScripting = true;
            elementListDictionary = new Dictionary<HtmlElement, List<HtmlElement>>();
            //elementLevelDictionary = new Dictionary<HtmlElement, int>();
            elementLevelDictionary = new Dictionary<int, List<HtmlElement>>();
            //elementListDictionary_MSHTML = new Dictionary<IHTMLElement, List<IHTMLElement>>();
            elementDicDictionary_MSHTML = new Dictionary<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>>();
            parentFrameIndexDictionary = new Dictionary<IHTMLElement, int>();
            elementLevelDictionary_MSHTML = new Dictionary<IHTMLElement, int>();
            buttonTargetDictionary = new Dictionary<Button, HtmlElement>();
            buttonTargetDictionary_MSHTML = new Dictionary<Button, IHTMLElement>();
            buttonTargetDictionary_Level = new Dictionary<Button, List<List<HtmlElement>>>();
            plexiglassDictionary  = new Dictionary<HtmlElement, Plexiglass>();
            plexiglassDictionary_MSHTML = new Dictionary<IHTMLElement, Plexiglass>();
            plexiglassDictionary_Level = new Dictionary<List<List<HtmlElement>>, Plexiglass>();
            detectedTable = new List<List<HtmlElement>>();
            detectedTable_MSHTML = new List<List<IHTMLElement>>();

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
            toolStripTextBox_URL.Width = this.Size.Width - 461;

            if(prevSize.Width - this.Size.Width != 0)
            {
                if (buttonTargetDictionary.Count + buttonTargetDictionary_MSHTML.Count + buttonTargetDictionary_Level.Count > 0)
                {
                    InitializeContents();
                }
            }
            prevSize = this.Size;
        }

        #endregion

        #region Initialization 메서드

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
                init=true;
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
            parentFrameIndexDictionary.Clear();
            elementLevelDictionary_MSHTML.Clear();
            detectedTable.Clear();
            detectedTable_MSHTML.Clear();
            foreach (Button btn in buttonTargetDictionary.Keys)
            {
                this.Controls.Remove(plexiglassDictionary[buttonTargetDictionary[btn]]);
                plexiglassDictionary[buttonTargetDictionary[btn]].Close();
                this.Controls.Remove(btn);
                btn.Dispose();
            }
            foreach (Button btn in buttonTargetDictionary_MSHTML.Keys)
            {
                this.Controls.Remove(plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[btn]]);
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[btn]].Close();
                this.Controls.Remove(btn);
                btn.Dispose();
            }
            foreach (Button btn in buttonTargetDictionary_Level.Keys)
            {
                this.Controls.Remove(plexiglassDictionary_Level[buttonTargetDictionary_Level[btn]]);
                plexiglassDictionary_Level[buttonTargetDictionary_Level[btn]].Close();
                this.Controls.Remove(btn);
                btn.Dispose();
            }
            buttonTargetDictionary.Clear();
            buttonTargetDictionary_MSHTML.Clear();
            buttonTargetDictionary_Level.Clear();
            plexiglassDictionary.Clear();
            plexiglassDictionary_MSHTML.Clear();
            plexiglassDictionary_Level.Clear();
            regionSelected = 0;

            this.ResumeLayout();
        }

        #endregion

        #region HTML Data Linkify

        private bool IsSeen(String style)
        {
            StyleGenerator sg = new StyleGenerator();
            sg.ParseStyleString(style == null ? "" : style);
            return !(sg.GetStyle("DISPLAY").Equals("none") || sg.GetStyle("VISIBILITY").Equals("hidden")
                                        || sg.GetStyle("WIDTH").Equals("0") || sg.GetStyle("HEIGHT").Equals("0")
                                        || sg.GetStyle("WIDTH").Equals("0px") || sg.GetStyle("HEIGHT").Equals("0px")
                                        || sg.GetStyle("FILTER").Contains("opacity=0") || sg.GetStyle("-ms-filter").Contains("opacity=0"));
        }

        private void FormLevelDictionary(HtmlElement htmlele, int level)
        {
            if (IsSeen(htmlele.Style) && !(htmlele.TagName.Equals("!")))
            {
                if (level != 0)
                {
                    HtmlElement parent = htmlele.Parent;
                    while (!parent.TagName.Equals("BODY"))
                    {
                        if (parent.TagName.Equals("TABLE") || parent.TagName.Equals("UL") || parent.TagName.Equals("OL") || parent.TagName.Equals("DL"))
                        {
                            return;
                        }
                        parent = parent.Parent;
                    }
                }

                if (!elementLevelDictionary.ContainsKey(level))
                {
                    List<HtmlElement> htmleleList = new List<HtmlElement>();
                    htmleleList.Add(htmlele);
                    elementLevelDictionary.Add(level, htmleleList);
                }
                else
                {
                    List<HtmlElement> htmleleList = elementLevelDictionary[level];
                    htmleleList.Add(htmlele);
                }
            }
        }

        void MakeStructure(HtmlElementCollection htmleleCollection)
        {
            int iframeIndex = 0;
            foreach (HtmlElement htmlele in htmleleCollection)
            {
                if (afterBody)
                {
                    if (htmlele.TagName.Equals("IFRAME"))
                    {
                        if (IsSeen(htmlele.Style))
                        {
                            HTMLDocument htmldoc = webBrowser.Document.DomDocument as HTMLDocument;
                            IHTMLWindow2 frame = (IHTMLWindow2)htmldoc.frames.item(iframeIndex);

                            try
                            {
                                HTMLDocument doc2 = (HTMLDocument)frame.document;
                                MakeStructure_MSHTML(htmlele.DomElement as IHTMLElement, doc2.documentElement.all, iframeIndex);
                            }
                            catch(System.UnauthorizedAccessException e)
                            {
                                e.ToString();
                            }
                        }
                        iframeIndex++;
                    }
                    else if ((!(htmlele.TagName.Equals("SCRIPT") || htmlele.TagName.Equals("STYLE") || htmlele.TagName.Equals("!"))) && IsSeen(htmlele.Style))
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

                    //elementLevelDictionary.Add(htmlele, 0);
                    FormLevelDictionary(htmlele, 0);

                    LinkElements(htmlele);
                }
                
            }
            afterBody = false;
        }

        void MakeStructure_MSHTML(IHTMLElement htmleleParent, IHTMLElementCollection htmleleCollection, int iframeIndex)
        {
            //int iframeIndex = 0;
            bool _afterBody = false;
            foreach (IHTMLElement htmlele_MSHTML in htmleleCollection)
            {
                if (_afterBody)
                {/*
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
                    else*/ if (!((htmlele_MSHTML.tagName.Equals("SCRIPT")) || (htmlele_MSHTML.tagName.Equals("STYLE")) || (htmlele_MSHTML.tagName.Equals("!"))))
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
                                    LinkElements_MSHTML(htmlele_MSHTML, htmleleParent, iframeIndex);
                                }
                            }
                            else
                            {
                                LinkElements_MSHTML(htmlele_MSHTML, htmleleParent, iframeIndex);
                            }
                        }
                        else
                        {
                            LinkElements_MSHTML(htmlele_MSHTML, htmleleParent, iframeIndex);
                        }
                    }
                }
                else if (htmlele_MSHTML.tagName.Equals("BODY"))
                {
                    _afterBody = true;

                    elementLevelDictionary_MSHTML.Add(htmlele_MSHTML, 0);

                    LinkElements_MSHTML(htmlele_MSHTML, htmleleParent, iframeIndex);
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
                    if (IsSeen(_htmlele.Style))
                        FormLevelAndElements(htmleleList, _htmlele);
                    /*if (_htmlele.Style != null)
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
                        
                        {
                            FormLevelAndElements(htmleleList, _htmlele);
                        }
                        else
                        {
                            FormLevelAndElements(htmleleList, _htmlele);
                        }
                    }
                    else
                    {
                        FormLevelAndElements(htmleleList, _htmlele);
                    }*/
                }
            }
            elementListDictionary.Add(htmlele, htmleleList);
        }

        private void LinkElements_MSHTML(IHTMLElement htmlele, IHTMLElement htmleleParent, int iframeindex)
        {
            List<IHTMLElement> htmleleList = new List<IHTMLElement>();
            foreach (IHTMLElement _htmlele in htmlele.children)
            {
                if (!((_htmlele.tagName.Equals("SCRIPT")) || (_htmlele.tagName.Equals("STYLE")) || (_htmlele.tagName.Equals("!"))))
                {
                    if (IsSeen(_htmlele.style.toString()))
                        FormLevelAndElements_MSHTML(htmleleList, _htmlele);
                    /*if (_htmlele.style != null)
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
                    }*/
                }
            }
            //elementListDictionary_MSHTML.Add(htmlele, htmleleList);

            Dictionary<IHTMLElement, List<IHTMLElement>> elementListDictionary_MSHTML;
            if (!elementDicDictionary_MSHTML.ContainsKey(htmleleParent))
            {
                elementListDictionary_MSHTML = new Dictionary<IHTMLElement, List<IHTMLElement>>();
                elementListDictionary_MSHTML.Add(htmlele, htmleleList);
                elementDicDictionary_MSHTML.Add(htmleleParent, elementListDictionary_MSHTML);
                parentFrameIndexDictionary.Add(htmleleParent, iframeindex);
            }
            else
            {
                elementListDictionary_MSHTML = elementDicDictionary_MSHTML[htmleleParent];
                elementListDictionary_MSHTML.Add(htmlele, htmleleList);
                elementDicDictionary_MSHTML[htmleleParent] = elementListDictionary_MSHTML;
                
            }
            if (!parentFrameIndexDictionary.ContainsKey(htmleleParent))
            {
                parentFrameIndexDictionary.Add(htmleleParent, iframeindex);
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
            /*if (elementLevelDictionary.ContainsKey(_htmlele))
                elementLevelDictionary.Remove(_htmlele);
            elementLevelDictionary.Add(_htmlele, level);*/
            FormLevelDictionary(_htmlele, level);
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

        private void toolStripButton_Detect_Click(object sender, EventArgs e)
        {
            if (init)
            {
                InitializeContents();
                InitializeView();
            }
            else
            {
                MessageBox.Show("웹페이지가 로드되지 않았습니다.");
            }
        }

        private void toolStripButton_Parse_Click(object sender, EventArgs e)
        {
            if (regionSelected == 0)
            {
                MessageBox.Show("파싱할 영역이 없습니다.");
                return;
            }
            ParseTargets();
        }

        private void toolStripButton_MultiPage_Click(object sender, EventArgs e)
        {
            if (mulForm == null)
            {
                mulForm = new MultiPageSettingForm(this);
                mulForm.Visible = true;
                mulForm.Show();
            }
            else if (mulForm.IsDisposed)
            {
                mulForm.Controls.Clear();
                mulForm = null;
                mulForm = new MultiPageSettingForm(this);
                mulForm.Visible = true;
                mulForm.Show();
            }
        }

        #endregion 

        #region Table Parse

        private void ParseTargets()
        {
            foreach (KeyValuePair<Button, HtmlElement> kvPair in buttonTargetDictionary.Where(kv => kv.Key.BackColor == Color.Aqua))
            {
                HtmlElement htmlele = kvPair.Value;
                List<HtmlElement> htmleleList = elementListDictionary[htmlele];

                if (htmlele.TagName.Equals("TABLE"))
                {
                    foreach (HtmlElement _htmlele in htmlele.Children)
                    {
                        if (!(_htmlele.TagName.Equals("COLGROUP") || _htmlele.TagName.Equals("ROWGROUP")))
                        {
                            if(IsSeen(_htmlele.Style))
                                FormCells(_htmlele);
                        }
                    }
                }
                else if (htmlele.TagName.Equals("UL") || htmlele.TagName.Equals("OL"))
                {
                    bool ok = true;
                    int count = 0;
                    foreach(HtmlElement htmlele_LI in htmlele.Children)
                    {
                        if(IsSeen(htmlele_LI.Style))
                        {
                            if (htmlele_LI.Children.Count == 1)
                            {
                                ok = false;
                                break;
                            }
                            /*
                            //TODO
                            foreach(HtmlElement cell in htmlele_LI.Children)
                            {
                                cell.ClientRectangle.Left;
                            }
                            */
                            count += htmlele_LI.Children.Count;
                        }
                    }
                    if(ok && (count > 4))
                        FormCells_ULOL(htmlele);
                }
                else if (htmlele.TagName.Equals("DL"))
                {
                    FormCells_DL(htmlele);
                }
            }

            foreach (KeyValuePair<Button, IHTMLElement> kvPair in buttonTargetDictionary_MSHTML.Where(kv => kv.Key.BackColor == Color.Aqua))
            {
                IHTMLElement htmlele = kvPair.Value;
                int level = elementLevelDictionary_MSHTML[htmlele];

                if (htmlele.tagName.Equals("TABLE"))
                {
                    foreach (IHTMLElement _htmlele in htmlele.children)
                    {
                        if (!(_htmlele.tagName.Equals("COLGROUP") || _htmlele.tagName.Equals("ROWGROUP")))
                        {
                            if (IsSeen(_htmlele.style.toString()))
                                FormCells_MSHTML(_htmlele);
                        }
                    }
                }
                else if (htmlele.tagName.Equals("UL") || htmlele.tagName.Equals("OL"))
                {
                    bool ok = true;
                    int count = 0;
                    foreach (IHTMLElement htmlele_LI in htmlele.children)
                    {
                        if (IsSeen(htmlele_LI.style.toString()))
                        {
                            if (htmlele_LI.children.Count == 1)
                            {
                                ok = false;
                                break;
                            }
                            /*
                            //TODO
                            foreach(HtmlElement cell in htmlele_LI.Children)
                            {
                                cell.ClientRectangle.Left;
                            }
                            */
                            count += htmlele_LI.children.Count;
                        }
                    }
                    if (ok && (count > 4))
                        FormCells_ULOL_MSHTML(htmlele);
                }
                else if (htmlele.tagName.Equals("DL"))
                {
                    FormCells_DL_MSHTML(htmlele);
                }
            }

            foreach(KeyValuePair<Button, List<List<HtmlElement>>> kvPair in buttonTargetDictionary_Level.Where(kv => kv.Key.BackColor == Color.Aqua))
            {
                FormCells_Level(kvPair.Value);
            }
        }

        int maxColumnCount_Table, maxRowCount_Table, selectedX = 0, selectedY = 0;

        private void FormCells(HtmlElement _htmlele, bool singlePage = true)
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
            if(singlePage)
                MoveToNextSheet();

            InsertDatas();
            detectedTable.Clear();
        }

        private void FormCells_ULOL(HtmlElement _htmlele, bool singlePage = true)
        {
            if (singlePage)
                MoveToNextSheet();

            InsertDatas_ULOL(_htmlele);
        }

        private void FormCells_DL(HtmlElement _htmlele, bool singlePage = true)
        {
            if (singlePage)
                MoveToNextSheet();

            InsertDatas_DL(_htmlele);
        }

        private void FormCells_ULOL_MSHTML(IHTMLElement _htmlele, bool singlePage = true)
        {
            if (singlePage)
                MoveToNextSheet();

            InsertDatas_ULOL_MSHTML(_htmlele);
        }

        private void FormCells_DL_MSHTML(IHTMLElement _htmlele, bool singlePage = true)
        {
            if (singlePage)
                MoveToNextSheet();

            InsertDatas_DL_MSHTML(_htmlele);
        }

        private void FormCells_MSHTML(IHTMLElement _htmlele, bool singlePage = true)
        {
            int count, maxCount = 0;
            maxRowCount_Table = 0;
            foreach (IHTMLElement row in _htmlele.children)
            {
                count = 0;
                List<IHTMLElement> htmleleList = new List<IHTMLElement>();
                foreach (IHTMLElement cell in row.children)
                {
                    htmleleList.Add(cell);
                    count++;
                }
                if (maxCount < count)
                    maxCount = count;
                if (htmleleList.Count != 0)
                {
                    detectedTable_MSHTML.Add(htmleleList);
                }
                maxRowCount_Table++;
            }
            maxColumnCount_Table = maxCount;
            if (singlePage)
                MoveToNextSheet();

            InsertDatas_MSHTML();
            detectedTable_MSHTML.Clear();
        }

        private void FormCells_Level(List<List<HtmlElement>> htmleleGroup, bool singlePage = true)
        {
            maxRowCount_Table = htmleleGroup.Count;
            maxColumnCount_Table = htmleleGroup[0].Count;
            if (singlePage)
                MoveToNextSheet();

            InsertDatas_Level(htmleleGroup);
        }

        private void InsertDatas()
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int endRow = detectedTable.Count;
            int endCol = GetEndColumn();
            int idxRow = 1, idxCol = 1;//, selectedX = activeCell.Row - 1, selectedY = activeCell.Column - 1;
            int rowspan = 1, colspan = 1;

            StyleGenerator sg = new StyleGenerator();
            Dictionary<int, Dictionary<int, int>> checkMatrix = new Dictionary<int, Dictionary<int, int>>();
            for (int i = 0; i < maxRowCount_Table; i++)
            {
                Dictionary<int, int> tempDic = new Dictionary<int, int>();
                for (int j = 0; j < maxColumnCount_Table; j++)
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
                    for (; idxCol <= maxColumnCount_Table; )
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
                        }
                    }

                    for (int i = idxRow - 1; i < (idxRow - 1 + rowspan); i++)
                    {
                        for (int j = idxCol - 1; j < (idxCol - 1 + colspan); j++)
                        {
                            checkMatrix[i][j] = 1;
                        }
                    }

                    String innerhtml = htmlele.InnerHtml == null ? "" : htmlele.InnerHtml;
                    innerhtml = OrganizeHtmls(innerhtml);

                    PutDataWithLink(activeSheet, idxRow, idxCol, innerhtml);

                    idxCol = idxCol + colspan;
                }
                idxRow++;
            }

            ResizeColumnsAndXY(activeSheet, idxRow, maxColumnCount_Table);
        }

        private void InsertDatas_ULOL(HtmlElement htmlele_ULOL)
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int idxRow = 1, idxCol = 1, maxIdxCol = 0;
            foreach (HtmlElement htmlele_LI in htmlele_ULOL.Children)
            {                
                idxCol = 1;
                if (IsSeen(htmlele_LI.Style))
                {
                    foreach (HtmlElement htmlele in htmlele_LI.Children)
                    {
                        if (IsSeen(htmlele.Style))
                        {
                            String innerhtml = htmlele.InnerHtml == null ? "" : htmlele.InnerHtml;
                            innerhtml = OrganizeHtmls(innerhtml);

                            PutDataWithLink(activeSheet, idxRow, idxCol, innerhtml);
                        }
                        idxCol = idxCol + 1;
                    }
                }
                if (maxIdxCol < (idxCol - 1))
                    maxIdxCol = idxCol - 1;
                idxRow++;
            }

            ResizeColumnsAndXY(activeSheet, idxRow, maxIdxCol);
        }

        private void InsertDatas_DL(HtmlElement htmlele_DL)
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int idxRow = 1, idxCol = 1;
            foreach (HtmlElement htmlele_D in htmlele_DL.Children)
            {
                idxCol = 1;
                if (IsSeen(htmlele_D.Style))
                {
                    String innerhtml = htmlele_D.InnerHtml == null ? "" : htmlele_D.InnerHtml;
                    innerhtml = OrganizeHtmls(innerhtml);

                    PutDataWithLink(activeSheet, idxRow, idxCol, innerhtml);
                    idxCol = idxCol + 1;
                }
                idxRow += (idxCol % 2);
            }

            ResizeColumnsAndXY(activeSheet, idxRow, 2);
        }

        private void InsertDatas_ULOL_MSHTML(IHTMLElement htmlele_ULOL)
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int idxRow = 1, idxCol = 1, maxIdxCol = 0;
            foreach (IHTMLElement htmlele_LI in htmlele_ULOL.children)
            {
                idxCol = 1;
                if (IsSeen(htmlele_LI.style.toString()))
                {
                    foreach (IHTMLElement htmlele in htmlele_LI.children)
                    {
                        if (IsSeen(htmlele.style.toString()))
                        {
                            String innerhtml = htmlele.innerHTML == null ? "" : htmlele.innerHTML;
                            innerhtml = OrganizeHtmls(innerhtml);

                            PutDataWithLink(activeSheet, idxRow, idxCol, innerhtml);
                        }
                        idxCol = idxCol + 1;
                    }
                }
                if (maxIdxCol < (idxCol - 1))
                    maxIdxCol = idxCol - 1;
                idxRow++;
            }

            ResizeColumnsAndXY(activeSheet, idxRow, maxIdxCol);
        }

        private void InsertDatas_DL_MSHTML(IHTMLElement htmlele_DL)
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int idxRow = 1, idxCol = 1;
            foreach (IHTMLElement htmlele_D in htmlele_DL.children)
            {
                idxCol = 1;
                if (IsSeen(htmlele_D.style.toString()))
                {
                    String innerhtml = htmlele_D.innerHTML == null ? "" : htmlele_D.innerHTML;
                    innerhtml = OrganizeHtmls(innerhtml);

                    int _idxCol = idxCol % 2;
                    PutDataWithLink(activeSheet, idxRow, _idxCol, innerhtml);

                    idxCol = idxCol + 1;
                }
                idxRow += (idxCol % 2);
            }

            ResizeColumnsAndXY(activeSheet, idxRow, 2);
        }

        private void InsertDatas_MSHTML()
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int endRow = detectedTable_MSHTML.Count;
            int endCol = GetEndColumn_MSHTML();
            int idxRow = 1, idxCol = 1;//, selectedX = activeCell.Row - 1, selectedY = activeCell.Column - 1;
            int rowspan = 1, colspan = 1, blank;

            StyleGenerator sg = new StyleGenerator();
            Dictionary<int, Dictionary<int, int>> checkMatrix = new Dictionary<int, Dictionary<int, int>>();
            for (int i = 0; i < maxRowCount_Table; i++)
            {
                Dictionary<int, int> tempDic = new Dictionary<int, int>();
                for (int j = 0; j < maxColumnCount_Table; j++)
                {
                    tempDic.Add(j, 0);
                }
                checkMatrix.Add(i, tempDic);
            }

            foreach (List<IHTMLElement> htmleleList in detectedTable_MSHTML)
            {
                idxCol = 1;
                blank = 0;
                foreach (IHTMLElement htmlele in htmleleList)
                {
                    rowspan = 1;
                    colspan = 1;

                    bool _skip = false;
                    for (; idxCol <= maxColumnCount_Table; )
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

                    sg.ParseStyleString(htmlele.style.toString() == null ? "" : htmlele.style.toString());

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
                        }
                    }

                    for (int i = idxRow - 1; i < (idxRow - 1 + rowspan); i++)
                    {
                        for (int j = idxCol - 1; j < (idxCol - 1 + colspan); j++)
                        {
                            checkMatrix[i][j] = 1;
                        }
                    }

                    String innerhtml = htmlele.innerHTML == null ? "" : htmlele.innerHTML;
                    innerhtml = OrganizeHtmls(innerhtml);

                    ((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]).Value2 = innerhtml;
                    Regex rg = new Regex("(링크 : (?<url>.*))");
                    Match m = rg.Match(innerhtml);
                    if (m.Groups["url"].Captures.Count == 1)
                    {
                        string url = m.Groups["url"].Value.EndsWith(")") ? m.Groups["url"].Value.Substring(0, m.Groups["url"].Value.Length - 1) : m.Groups["url"].Value;
                        IHTMLElement parent = null;
                        foreach (KeyValuePair<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>> kv in elementDicDictionary_MSHTML)
                        {
                            if (kv.Value.ContainsKey(htmlele))
                            {
                                parent = kv.Key;
                                break;
                            }
                        }
                        Uri uri = new Uri(parent.getAttribute("src"));

                        if (url.StartsWith("http://") || url.StartsWith("https://"))
                            activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]), url);
                        else if (url.StartsWith("./"))
                            activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                                uri.GetLeftPart(UriPartial.Path) + url);
                        else if (url.StartsWith("../"))
                            activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                                uri.GetLeftPart(UriPartial.Authority) + url);
                        else if (url.StartsWith(".../"))
                            activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                                uri.Authority + url);
                        else
                        {
                            if (url.StartsWith("/"))
                            {
                                activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                                    uri.GetLeftPart(UriPartial.Authority) + url);
                            }
                            else
                            {
                                activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                                    uri.GetLeftPart(UriPartial.Path) + url);
                            }
                        }
                    }

                    idxCol = idxCol + colspan;
                }

                for (int i = idxRow; i <= (idxRow - 1 + rowspan); i++)
                {
                    for (int j = 1; j <= maxColumnCount_Table; j++)
                    {
                        object target = ((Excel.Range)activeSheet.Cells[i + selectedX, j + selectedY]).Value2 == null
                            ? "" : ((Excel.Range)activeSheet.Cells[i + selectedX, j + selectedY]).Value2;

                        if (Regex.Replace(target.ToString(),
                            "\n", string.Empty).Trim().Length == 0)
                            blank++;
                    }
                }
                if (blank == maxColumnCount_Table)
                {
                    for (int i = idxRow - 1; i < (idxRow - 1 + rowspan); i++)
                    {
                        for (int j = 0; j < maxColumnCount_Table; j++)
                        {
                            checkMatrix[i][j] = 0;
                        }
                    }
                }
                else
                {
                    idxRow++;
                }
            }

            ResizeColumnsAndXY(activeSheet, idxRow, maxColumnCount_Table);
        }

        private void ResizeColumnsAndXY(Excel.Worksheet activeSheet, int idxRow, int maxIdxCol)
        {
            for (int i = 1; i <= maxColumnCount_Table; i++)
            {
                activeSheet.Columns[1][i].AutoFit();
                if (activeSheet.Columns[1][i].ColumnWidth > 30)
                    activeSheet.Columns[1][i].ColumnWidth = 30; // activeSheet.Columns.EntireColumn.AutoFit();
            }
            selectedX = idxRow + selectedX + 1;
            selectedY = 0;
        }

        private void InsertDatas_Level(List<List<HtmlElement>> htmleleGroup)
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet activeSheet = addins.GetActiveWorksheet();
            Excel.Range activeCell = (Excel.Range)addins.Application.ActiveCell;

            int idxRow = 1, idxCol = 1, maxIdxCol = 0;
            foreach (List<HtmlElement> htmleleList in htmleleGroup)
            {
                idxCol = 1;
                foreach (HtmlElement htmlele in htmleleList)
                {
                    String innerhtml = htmlele.InnerHtml == null ? "" : htmlele.InnerHtml;
                    innerhtml = OrganizeHtmls(innerhtml);

                    PutDataWithLink(activeSheet, idxRow, idxCol, innerhtml);

                    idxCol = idxCol + 1;
                }
                if (maxIdxCol < (idxCol - 1))
                    maxIdxCol = idxCol - 1;
                idxRow++;
            }

            ResizeColumnsAndXY(activeSheet, idxRow, maxIdxCol);
        }

        private void PutDataWithLink(Excel.Worksheet activeSheet, int idxRow, int idxCol, String innerhtml)
        {
            ((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]).Value2 = innerhtml;
            Regex rg = new Regex("(링크 : (?<url>.*))");
            Match m = rg.Match(innerhtml);
            if (m.Groups["url"].Captures.Count == 1)
            {
                string url = m.Groups["url"].Value.EndsWith(")") ? m.Groups["url"].Value.Substring(0, m.Groups["url"].Value.Length - 1) : m.Groups["url"].Value;
                if (url.StartsWith("http://") || url.StartsWith("https://"))
                    activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]), url);
                else if (url.StartsWith("./"))
                    activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                        webBrowser.Url.GetLeftPart(UriPartial.Path) + url);
                else if (url.StartsWith("../"))
                    activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                        webBrowser.Url.GetLeftPart(UriPartial.Authority) + url);
                else if (url.StartsWith(".../"))
                    activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                        webBrowser.Url.Authority + url);
                else
                {
                    if (url.StartsWith("/"))
                    {
                        activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                            webBrowser.Url.GetLeftPart(UriPartial.Authority) + url);
                    }
                    else
                    {
                        activeSheet.Hyperlinks.Add(((Excel.Range)activeSheet.Cells[idxRow + selectedX, idxCol + selectedY]),
                            webBrowser.Url.GetLeftPart(UriPartial.Path) + url);
                    }
                }
            }
        }

        private String OrganizeHtmls(String innerhtml)
        {
            innerhtml = Regex.Replace(innerhtml, "<style(.|\n)*?</style>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<script(.|\n)*?</script>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<!--(.|\n)*?-->", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<div(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</div>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<span(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</span>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<p(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</p>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<em(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</em>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<strong(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</strong>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<caption(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</caption>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<h1(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</h1>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<h2(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</h2>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<h3(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</h3>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<h4(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</h4>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<h5(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</h5>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<h6(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</h6>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<label(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "</label>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<hr(.|\n)*?>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<br(.|\n)*?/>", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<a(?<str1>.|\n)*?href=\"(?<str2>.*?)\"(?<str3>.|\n)*?>(?<str4>.*?)</a>", "${str4}\n(링크 : ${str2})");
            innerhtml = Regex.Replace(innerhtml, "\t", string.Empty);
            innerhtml = Regex.Replace(innerhtml, "<(?<str1>.*?) style=\"(?<str2>.*?)display:(?<str3>.*?)none;(?<str4>.*?)\"(?<str5>.*?)>(?<str6>.*?)<(?<str7>.*?)>", "${str1}" == "${str7}" ? string.Empty : "<${str1} style=\"${str2}display:${str3}none;${str4}\"${str5}>${str6}<${str7}>");
            innerhtml = Regex.Replace(innerhtml, "&amp;", "&");
            innerhtml = Regex.Replace(innerhtml, "&quot;", "\"");
            innerhtml = Regex.Replace(innerhtml, "&lt;", "<");
            innerhtml = Regex.Replace(innerhtml, "&gt;", ">");
            innerhtml = Regex.Replace(innerhtml, @"\n\s+", "\n");
            innerhtml = Regex.Replace(innerhtml, @"\n", "\n");
            innerhtml = Regex.Replace(innerhtml, @"\r\n\s+", "\r\n");
            innerhtml = Regex.Replace(innerhtml, @"\r\n", "\r\n");

            return innerhtml;
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

        private int GetEndColumn_MSHTML()
        {
            int maxCol = 0;
            foreach (List<IHTMLElement> htmleleList in detectedTable_MSHTML)
            {
                if (maxCol < htmleleList.Count)
                    maxCol = htmleleList.Count;
            }
            return maxCol;
        }

        #endregion

        #region MultiplePage Parse

        public String frontURL
        {
            get;
            set;
        }
        public String backURL
        {
            get;
            set;
        }
        public int startIndex
        {
            get;
            set;
        }
        public int endIndex
        {
            get;
            set;
        }
        public bool working = false;

        public Dictionary<HtmlElement, List<List<String>>> targetTags;
        public Dictionary<IHTMLElement, List<List<String>>> targetTags_MSHTML;
        public int failCount, pageIndex;
        public String pageURL;

        public void MultiPageParse()
        {
            if (regionSelected == 0)
            {
                MessageBox.Show("파싱할 영역이 없습니다.");
                return;
            }

            progressValue = 0;
            if (backgroundWorker_Init.IsBusy != true)
            {
                alert_Init = new AlertForm();
                alert_Init.Show();
                backgroundWorker_Init.RunWorkerAsync();
            }

            failCount = 0;
            WebBrowser wb = new WebBrowser();
            wb.ScriptErrorsSuppressed = true;
            HtmlDocument htmldoc = wb.Document;
            targetTags = new Dictionary<HtmlElement, List<List<string>>>();
            foreach (KeyValuePair<Button, HtmlElement> kvPair in buttonTargetDictionary.Where(kv => kv.Key.BackColor == Color.Aqua))
            {
                HtmlElement htmlele = kvPair.Value;

                List<List<String>> tagAnalytics = new List<List<string>>();
                if (htmlele.Id == null)
                {
                    tagAnalytics = AnalyzeTarget(htmlele);
                }
                targetTags.Add(htmlele, tagAnalytics);   
            }

            HTMLDocument doc = webBrowser.Document.DomDocument as HTMLDocument;
            targetTags_MSHTML = new Dictionary<IHTMLElement, List<List<string>>>();
            foreach (KeyValuePair<Button, IHTMLElement> kvPair in buttonTargetDictionary_MSHTML.Where(kv => kv.Key.BackColor == Color.Aqua))
            {
                IHTMLElement htmlele_MSHTML = kvPair.Value;
                List<List<String>> tagAnalytics = null;
                if (htmlele_MSHTML.id == null)
                {
                    tagAnalytics = AnalyzeTarget_MSHTML(htmlele_MSHTML);
                }
                targetTags_MSHTML.Add(htmlele_MSHTML, tagAnalytics);
            }
            progressValue = 10;

            pageIndex = startIndex;
            pageURL = frontURL + pageIndex + backURL;
            wb.DocumentCompleted += wb_DocumentCompleted;
            wb.Navigate(pageURL);
        }
        
        private void wb_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser wb = sender as WebBrowser;
            if (wb.ReadyState == WebBrowserReadyState.Complete)
            {
                MoveToNextSheet();

                ParseEachPage(wb);
                ParseEachPage_MSHTML();

                if ((progressValue + (80 / (endIndex - startIndex + 1))) <= 90)
                    progressValue += (80 / (endIndex - startIndex + 1));

                InvokeNextPage(wb);
            }
        }

        private void MoveToNextSheet()
        {
            var addins = Globals.ThisAddIn;
            Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)addins.Application.Worksheets.Add();
            selectedX = 0; selectedY = 0;
        }

        private void InvokeNextPage(WebBrowser wb)
        {
            if (pageIndex <= endIndex)
            {
                pageIndex++;
                pageURL = frontURL + pageIndex + backURL;
                wb.Navigate(pageURL);
            }
            else
            {
                InitializeContents();
                progressValue = 100;
            }
        }

        private HtmlElement GetSameElement(Dictionary<HtmlElement, List<List<String>>> fromTags, List<List<String>> tags, HtmlElement sameEle)
        {
            foreach (KeyValuePair<HtmlElement, List<List<String>>> _kv in fromTags)
            {
                HtmlElement _htmlele = _kv.Key;
                List<List<String>> _tags = _kv.Value;
                int i = 0;
                bool notEqual = false;
                if (tags.Count == _tags.Count)
                {
                    foreach (List<String> tdTag in tags)
                    {
                        List<String> _tdTag = _tags[i];
                        if (tdTag.Count == _tdTag.Count)
                        {
                            int j = 0;
                            foreach (String eachTag in tdTag)
                            {
                                String _eachTag = _tdTag[j];
                                if (!eachTag.Equals(_eachTag))
                                {
                                    notEqual = true;
                                    break;
                                }
                                j++;
                            }
                            if (notEqual)
                                break;
                        }
                        else
                        {
                            notEqual = true;
                            break;
                        }
                        i++;
                    }
                }
                else
                    continue;

                if (!notEqual)
                {
                    sameEle = _htmlele;
                    break;
                }
            }
            return sameEle;
        }

        private void ParseEachPage(WebBrowser wb)
        {
            HtmlDocument htmldoc = wb.Document;
            Dictionary<HtmlElement, List<List<String>>> fromTags = new Dictionary<HtmlElement, List<List<string>>>();
            foreach (HtmlElement htmlele_WB in wb.Document.All)
            {
                if (IsSeen(htmlele_WB.Style) && (htmlele_WB.TagName.Equals("TABLE") || htmlele_WB.TagName.Equals("UL")
                    || htmlele_WB.TagName.Equals("OL") || htmlele_WB.TagName.Equals("DL")))
                {
                    fromTags.Clear();

                    if (htmlele_WB.Id == null)
                    {
                        fromTags.Add(htmlele_WB, AnalyzeTarget(htmlele_WB));
                    }

                    string tagname = htmlele_WB.TagName;
                    foreach (KeyValuePair<Button, HtmlElement> kvPair
                        in buttonTargetDictionary.Where(kv => kv.Value.TagName.Equals(tagname) && kv.Key.BackColor == Color.Aqua))
                    {
                        HtmlElement htmlele = kvPair.Value;
                        if (htmlele.Id != null)
                        {
                            if (tagname.Equals("TABLE"))
                                FormCells(htmldoc.GetElementById(htmlele.Id), false);
                            else if (tagname.Equals("UL") || tagname.Equals("OL"))
                                FormCells_ULOL(htmldoc.GetElementById(htmlele.Id), false);
                            else
                                FormCells_DL(htmldoc.GetElementById(htmlele.Id), false);
                        }
                        else if (targetTags.ContainsKey(htmlele))
                        {
                            List<List<String>> tags = targetTags[htmlele];
                            HtmlElement sameEle = null;
                            sameEle = GetSameElement(fromTags, tags, sameEle);

                            if (sameEle != null)
                            {
                                if (tagname.Equals("TABLE"))
                                {
                                    foreach (HtmlElement resultele in sameEle.Children)
                                    {
                                        if ((!IsSeen(resultele.Style))
                                            && ((resultele.TagName.Equals("THEAD")) || (resultele.TagName.Equals("TBODY")) || (resultele.TagName.Equals("TFOOT"))))
                                        {
                                            FormCells(resultele, false);
                                        }
                                    }
                                }
                                else if (tagname.Equals("UL") || tagname.Equals("OL"))
                                    FormCells_ULOL(sameEle, false);
                                else
                                    FormCells_DL(sameEle, false);
                            }
                            else
                            {
                                failCount++;
                            }
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                }
            }
        }

        private void ParseEachPage_MSHTML()
        {
            HTMLDocument doc = webBrowser.Document.DomDocument as HTMLDocument;
            foreach (KeyValuePair<Button, IHTMLElement> kvPair in buttonTargetDictionary_MSHTML.Where(kv => kv.Key.BackColor == Color.Aqua))
            {
                IHTMLElement htmlele_MSHTML = kvPair.Value;
                IHTMLElement parent = null;
                foreach (KeyValuePair<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>> kv in elementDicDictionary_MSHTML)
                {
                    if (kv.Value.ContainsKey(htmlele_MSHTML))
                    {
                        parent = kv.Key;
                        break;
                    }
                }
                int iframeIndex = parentFrameIndexDictionary[parent];
                IHTMLWindow2 frame = (IHTMLWindow2)doc.frames.item(iframeIndex);

                try
                {
                    HTMLDocument htmldoc_frame = (HTMLDocument)frame.document;

                    string tagname = htmlele_MSHTML.tagName;
                    if (htmlele_MSHTML.id != null)
                    {
                        if (tagname.Equals("TABLE"))
                            FormCells_MSHTML(htmldoc_frame.getElementById(htmlele_MSHTML.id), false);
                        else if (tagname.Equals("UL") || tagname.Equals("OL"))
                            FormCells_ULOL_MSHTML(htmldoc_frame.getElementById(htmlele_MSHTML.id), false);
                        else
                            FormCells_DL_MSHTML(htmldoc_frame.getElementById(htmlele_MSHTML.id), false);
                        continue;
                    }
                    else if (targetTags_MSHTML.ContainsKey(htmlele_MSHTML))
                    {
                        Dictionary<IHTMLElement, List<List<String>>> fromTags_MSHTML = new Dictionary<IHTMLElement, List<List<String>>>();
                        foreach (IHTMLElement htmlele_Frame in htmldoc_frame.all)
                        {
                            if (htmlele_Frame.tagName.Equals(tagname))
                            {
                                fromTags_MSHTML.Clear();
                                if (IsSeen(htmlele_Frame.style.toString()))
                                {
                                    if (htmlele_Frame.id == null)
                                    {
                                        fromTags_MSHTML.Add(htmlele_Frame, AnalyzeTarget_MSHTML(htmlele_Frame));
                                    }
                                }

                                List<List<String>> tags = targetTags_MSHTML[htmlele_MSHTML];
                                IHTMLElement sameEle = null;
                                foreach (KeyValuePair<IHTMLElement, List<List<String>>> _kv in fromTags_MSHTML)
                                {
                                    IHTMLElement _htmlele = _kv.Key;
                                    List<List<String>> _tags = _kv.Value;
                                    int i = 0;
                                    bool notEqual = false;
                                    if (tags.Count == _tags.Count)
                                    {
                                        foreach (List<String> tdTag in tags)
                                        {
                                            List<String> _tdTag = _tags[i];
                                            if (tdTag.Count == _tdTag.Count)
                                            {
                                                int j = 0;
                                                foreach (String eachTag in tdTag)
                                                {
                                                    String _eachTag = _tdTag[j];
                                                    if (!eachTag.Equals(_eachTag))
                                                    {
                                                        notEqual = true;
                                                        break;
                                                    }
                                                    j++;
                                                }
                                                if (notEqual)
                                                    break;
                                            }
                                            else
                                            {
                                                notEqual = true;
                                                break;
                                            }
                                            i++;
                                        }
                                    }
                                    else
                                        continue;

                                    if (!notEqual)
                                    {
                                        sameEle = _htmlele;
                                        break;
                                    }
                                }


                                if (sameEle != null)
                                {
                                    if (tagname.Equals("TABLE"))
                                    {
                                        foreach (IHTMLElement resultele in sameEle.children)
                                        {
                                            if (IsSeen(resultele.style.toString())
                                                && ((resultele.tagName.Equals("THEAD")) || (resultele.tagName.Equals("TBODY")) || (resultele.tagName.Equals("TFOOT"))))
                                            {
                                                FormCells_MSHTML(resultele, false);
                                            }
                                        }
                                    }
                                    else if (tagname.Equals("UL") || tagname.Equals("OL"))
                                        FormCells_ULOL_MSHTML(sameEle, false);
                                    else
                                        FormCells_DL_MSHTML(sameEle, false);
                                }
                                else
                                {
                                    failCount++;
                                }
                            }
                        }
                    }
                    else
                    {
                        failCount++;
                    }
                }
                catch
                {

                }
            }
        }

        private List<List<String>> AnalyzeTarget(HtmlElement htmlele)
        {
            List<List<String>> tagAnalytics = new List<List<string>>();
            if (htmlele.TagName.Equals("TABLE"))
            {
                foreach (HtmlElement htmlele_upper in htmlele.Children)
                {
                    if (IsSeen(htmlele_upper.Style)
                        && (htmlele_upper.TagName.Equals("THEAD") || htmlele_upper.TagName.Equals("TBODY") || htmlele_upper.TagName.Equals("TFOOT")))
                    {
                        foreach (HtmlElement htmlele_tr in htmlele.Children)
                        {
                            foreach (HtmlElement htmlele_td in htmlele_tr.Children)
                            {
                                List<String> tags = new List<String>();
                                foreach (HtmlElement htmlele_tdChild in htmlele_td.Children)
                                {
                                    if (IsSeen(htmlele_tdChild.Style))
                                    {
                                        tags.Add(htmlele_tdChild.TagName);
                                    }
                                }
                                tagAnalytics.Add(tags);
                            }
                        }
                    }
                }
            }
            else if(htmlele.TagName.Equals("UL") || htmlele.TagName.Equals("OL"))
            {
                foreach (HtmlElement htmlele_LI in htmlele.Children)
                {
                    if (IsSeen(htmlele_LI.Style))
                    {
                        List<String> tags = new List<String>();
                        foreach (HtmlElement htmlele_Each in htmlele.Children)
                        {
                            if (IsSeen(htmlele_Each.Style))
                            {
                                tags.Add(htmlele_Each.TagName);
                            }
                        }
                        tagAnalytics.Add(tags);
                    }
                }
            }
            else if (htmlele.TagName.Equals("DL"))
            {
                foreach (HtmlElement htmlele_DTDD in htmlele.Children)
                {
                    if (IsSeen(htmlele_DTDD.Style))
                    {
                        List<String> tags = new List<String>();
                        foreach (HtmlElement htmlele_Each in htmlele_DTDD.Children)
                        {
                            if (IsSeen(htmlele_Each.Style))
                            {
                                tags.Add(htmlele_Each.TagName);
                            }
                        }
                        tagAnalytics.Add(tags);
                    }
                }
            }

            return tagAnalytics;
        }

        private List<List<String>> AnalyzeTarget_MSHTML(IHTMLElement htmlele_MSHTML)
        {
            List<List<String>> tagAnalytics = new List<List<string>>();
            if (htmlele_MSHTML.tagName.Equals("TABLE"))
            {
                foreach (IHTMLElement htmlele_upper in htmlele_MSHTML.children)
                {
                    if (IsSeen(htmlele_upper.style.toString())
                        && (htmlele_upper.tagName.Equals("THEAD") || htmlele_upper.tagName.Equals("TBODY") || htmlele_upper.tagName.Equals("TFOOT")))
                    {
                        foreach (IHTMLElement htmlele_tr in htmlele_MSHTML.children)
                        {
                            foreach (IHTMLElement htmlele_td in htmlele_tr.children)
                            {
                                List<String> tags = new List<String>();
                                foreach (IHTMLElement htmlele_tdChild in htmlele_td.children)
                                {
                                    if (IsSeen(htmlele_tdChild.style.toString()))
                                    {
                                        tags.Add(htmlele_tdChild.tagName);
                                    }
                                }
                                tagAnalytics.Add(tags);
                            }
                        }
                    }
                }
            }
            else if (htmlele_MSHTML.tagName.Equals("UL") || htmlele_MSHTML.tagName.Equals("OL"))
            {
                foreach (IHTMLElement htmlele_MSHTML_LI in htmlele_MSHTML.children)
                {
                    if (IsSeen(htmlele_MSHTML_LI.style.toString()))
                    {
                        List<String> tags = new List<String>();
                        foreach (IHTMLElement htmlele_MSHTML_Each in htmlele_MSHTML.children)
                        {
                            if (IsSeen(htmlele_MSHTML_Each.style.toString()))
                            {
                                tags.Add(htmlele_MSHTML_Each.tagName);
                            }
                        }
                        tagAnalytics.Add(tags);
                    }
                }
            }
            else if (htmlele_MSHTML.tagName.Equals("DL"))
            {
                foreach (IHTMLElement htmlele_MSHTML_THTD in htmlele_MSHTML.children)
                {
                    if (IsSeen(htmlele_MSHTML_THTD.style.toString()))
                    {
                        List<String> tags = new List<String>();
                        foreach (IHTMLElement htmlele_MSHTML_Each in htmlele_MSHTML_THTD.children)
                        {
                            if (IsSeen(htmlele_MSHTML_Each.style.toString()))
                            {
                                tags.Add(htmlele_MSHTML_Each.tagName);
                            }
                        }
                        tagAnalytics.Add(tags);
                    }
                }
            }

            return tagAnalytics;
        }

        #endregion

        #region webBroswer 객체용 메서드

        public void OnScrollEventHandler(object sender, EventArgs e)
        {
            int scrollTop = webBrowser.Document.GetElementsByTagName("HTML")[0].ScrollTop,
                scrollLeft = webBrowser.Document.GetElementsByTagName("HTML")[0].ScrollLeft;
            int xDiff = docLocation.X - scrollLeft, yDiff = docLocation.Y - scrollTop;
            if ((Math.Abs(xDiff) >= 10) || (Math.Abs(yDiff) >= 10))
            {
                docLocation.X = scrollLeft;
                docLocation.Y = scrollTop;
                Point parentPtr = this.PointToScreen(Point.Empty);

                ScrollControls(xDiff, yDiff, parentPtr);
                ScrollControl_MSHTML(xDiff, yDiff, parentPtr);
                ScrollControls_Level(xDiff, yDiff, parentPtr);

                this.ResumeLayout();
                this.ActiveControl = webBrowser;
            }
        }

        private void ScrollControls(int xDiff, int yDiff, Point parentPtr)
        {
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
            
            foreach (KeyValuePair<HtmlElement, Plexiglass> kvPair in plexiglassDictionary)
            {
                Plexiglass plexiglass = kvPair.Value;
                HtmlElement htmlele = kvPair.Key;

                int x = plexiglass.Location.X, y = plexiglass.Location.Y,
                    width = plexiglass.Width, height = plexiglass.Height;
                if (Math.Abs(xDiff) > 0)
                {
                    if ((plexiglass._offsetLeft > docLocation.X) && (plexiglass._offsetLeft < webBrowser.Width + docLocation.X))
                    {
                        x = plexiglass._offsetLeft - (docLocation.X);
                        plexiglass.Location = new Point(x + parentPtr.X, y);
                        if ((parentPtr.X + this.Size.Width - (plexiglass.Location.X + plexiglass._width)) > 0)
                        {
                            plexiglass.ClientSize = new Size((plexiglass._width * 2) - parentPtr.X - this.Size.Width + x, plexiglass.ClientSize.Height);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass._width, plexiglass.ClientSize.Height);
                        }
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

                        if ((parentPtr.X + this.Size.Width - (plexiglass.Location.X + plexiglass._width)) < 0)
                        {
                            plexiglass.ClientSize = new Size((plexiglass._width * 2) - parentPtr.X - this.Size.Width + x, plexiglass.ClientSize.Height);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass._width, plexiglass.ClientSize.Height);
                        }
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

                        if (plexiglass.Location.Y + plexiglass.Height > parentPtr.Y + this.Size.Height)
                        {
                            plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, this.Size.Height - y - 50);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, plexiglass._height);
                        }
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
        }

        private void ScrollControl_MSHTML(int xDiff, int yDiff, Point parentPtr)
        {
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

            foreach (KeyValuePair<IHTMLElement, Plexiglass> kvPair in plexiglassDictionary_MSHTML)
            {
                Plexiglass plexiglass = kvPair.Value;
                IHTMLElement htmlele_MSHTML = kvPair.Key;

                int x = plexiglass.Location.X, y = plexiglass.Location.Y,
                    width = plexiglass.Width, height = plexiglass.Height;
                if (Math.Abs(xDiff) > 0)
                {
                    if ((plexiglass._offsetLeft > docLocation.X) && (plexiglass._offsetLeft < webBrowser.Width + docLocation.X))
                    {
                        x = plexiglass._offsetLeft - (docLocation.X);
                        plexiglass.Location = new Point(x + parentPtr.X, y);
                        if (plexiglass.Location.X + plexiglass._width > parentPtr.X + this.Size.Width - 50)
                        {
                            plexiglass.ClientSize = new Size(this.Size.Width - y - 50, plexiglass.ClientSize.Height);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass._width, plexiglass.ClientSize.Height);
                        }
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
                        //y = plexiglass.Location.Y - parentPtr.Y + yDiff;
                        y = plexiglass._offsetTop + 33 - docLocation.Y;
                        plexiglass.Location = new Point(x, y + parentPtr.Y);
                        if (plexiglass.Location.Y + plexiglass._height > parentPtr.Y + this.Size.Height - 50)
                        {
                            plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, this.Size.Height - y - 50);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, plexiglass._height);
                        }
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
        }

        private void ScrollControls_Level(int xDiff, int yDiff, Point parentPtr)
        {
            foreach (Button btn in buttonTargetDictionary_Level.Keys)
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

            foreach (KeyValuePair<List<List<HtmlElement>>, Plexiglass> kvPair in plexiglassDictionary_Level)
            {
                Plexiglass plexiglass = kvPair.Value;
                List<List<HtmlElement>> htmleleGroup = kvPair.Key;

                int x = plexiglass.Location.X, y = plexiglass.Location.Y,
                    width = plexiglass.Width, height = plexiglass.Height;
                if (Math.Abs(xDiff) > 0)
                {
                    if ((plexiglass._offsetLeft > docLocation.X) && (plexiglass._offsetLeft < webBrowser.Width + docLocation.X))
                    {
                        x = plexiglass._offsetLeft - (docLocation.X);
                        plexiglass.Location = new Point(x + parentPtr.X, y);
                        if ((parentPtr.X + this.Size.Width - (plexiglass.Location.X + plexiglass._width)) > 0)
                        {
                            plexiglass.ClientSize = new Size((plexiglass._width * 2) - parentPtr.X - this.Size.Width + x, plexiglass.ClientSize.Height);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass._width, plexiglass.ClientSize.Height);
                        }
                        if (plexiglass._offsetLeft - docLocation.X > webBrowser.Width)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary_Level.FirstOrDefault(kv => kv.Value == htmleleGroup).Key.BackColor == Color.Aqua)
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

                        if ((parentPtr.X + this.Size.Width - (plexiglass.Location.X + plexiglass._width)) < 0)
                        {
                            plexiglass.ClientSize = new Size((plexiglass._width * 2) - parentPtr.X - this.Size.Width + x, plexiglass.ClientSize.Height);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass._width, plexiglass.ClientSize.Height);
                        }
                        if (applyWidth <= 0)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary_Level.FirstOrDefault(kv => kv.Value == htmleleGroup).Key.BackColor == Color.Aqua)
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

                        if (plexiglass.Location.Y + plexiglass.Height > parentPtr.Y + this.Size.Height)
                        {
                            plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, this.Size.Height - y - 50);
                        }
                        else
                        {
                            plexiglass.ClientSize = new Size(plexiglass.ClientSize.Width, plexiglass._height);
                        }
                        if (plexiglass._offsetTop - docLocation.Y > webBrowser.Height)
                            plexiglass.Visible = false;
                        else
                        {
                            if (buttonTargetDictionary_Level.FirstOrDefault(kv => kv.Value == htmleleGroup).Key.BackColor == Color.Aqua)
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
                            if (buttonTargetDictionary_Level.FirstOrDefault(kv => kv.Value == htmleleGroup).Key.BackColor == Color.Aqua)
                                plexiglass.Visible = true;
                            else
                                plexiglass.Visible = false;
                        }
                    }
                }
            }
        }

        #endregion

        #region 영역 및 선택 버튼 생성 함수

        private void FormButtons()
        {
            CatchTables();
            CatchTables_MSHTML();

            CatchTablesByLevel();

            this.ResumeLayout();
        }

        private void CatchTables()
        {
            foreach (HtmlElement htmlele in elementListDictionary.Keys)
            {
                if (IsSeen(htmlele.Style))
                {
                    if(htmlele.TagName.Equals("TABLE") || htmlele.TagName.Equals("UL") || htmlele.TagName.Equals("OL") || htmlele.TagName.Equals("DL"))
                    {
                        CreateButton(htmlele);
                    }
                }
            }
        }

        private void CatchTables_MSHTML()
        {
            foreach (KeyValuePair<IHTMLElement, Dictionary<IHTMLElement, List<IHTMLElement>>> kvPair in elementDicDictionary_MSHTML)
            {
                IHTMLElement htmleleParent = kvPair.Key;
                Dictionary<IHTMLElement, List<IHTMLElement>> elementListDictionary_MSHTML = kvPair.Value;
                foreach (IHTMLElement htmlele_MSHTML in elementListDictionary_MSHTML.Keys)
                {
                    if (IsSeen(htmlele_MSHTML.style.toString()) && htmlele_MSHTML.tagName.Equals("TABLE"))
                    {
                        CreateButton_MSHTML(htmlele_MSHTML, htmleleParent);
                    }
                }
            }
        }

        private void CatchTablesByLevel()
        {
            int rowCount = 0, rowGroupCount = 0, colCount = 0, colGroupCount = 0;
            List<List<List<HtmlElement>>> rowGroups = new List<List<List<HtmlElement>>>();
            foreach (KeyValuePair<int, List<HtmlElement>> kvPair in elementLevelDictionary)
            {
                int level = kvPair.Key;
                List<HtmlElement> htmleleList = kvPair.Value;

                ////Dictionary<int, Dictionary<int, HtmlElement>> matrix = new Dictionary<int, Dictionary<int, HtmlElement>>();
                Dictionary<int, Dictionary<HtmlElement, int>> matrix = new Dictionary<int, Dictionary<HtmlElement, int>>();
                foreach (HtmlElement htmlele in htmleleList)
                {
                    int x, y;
                    SetElementXY(htmlele, out x, out y);
                    if (!matrix.ContainsKey(y))
                    {
                        Dictionary<HtmlElement, int> dic = new Dictionary<HtmlElement, int>();
                        dic.Add(htmlele, x);
                        matrix.Add(y, dic);
                    }
                    else
                    {
                        Dictionary<HtmlElement, int> dic = matrix[y];
                        dic.Add(htmlele, x);
                    }
                }

                List<List<HtmlElement>> rowGroup = new List<List<HtmlElement>>();
                HtmlElement prevele = null;
                int prev_x = -1, prev_y = -1, diffX = 0, diffY = 0, prevHeight = 0, initY = -1, matrixCount = 0, matrixCount_prev = -1;

                foreach (Dictionary<HtmlElement, int> dic in matrix.Values)
                    matrixCount += dic.Count;
                //for(int i = 0; i < this.Size.Width; i++)
                while ((matrixCount != 0) && (matrixCount != matrixCount_prev))
                {
                    matrixCount_prev = matrixCount;
                    
                    foreach (int y in matrix.Keys.OrderBy(_y => _y).ToList())
                    {
                        colCount = 0;
                        List<HtmlElement> row = new List<HtmlElement>();
                        List<HtmlElement> rowXs = new List<HtmlElement>();
                        foreach (KeyValuePair<HtmlElement, int> _kvPair in matrix[y].OrderBy(_x => _x.Value).ToList())
                        {
                            HtmlElement htmlele = _kvPair.Key;
                            int x = _kvPair.Value;
                            if (colCount == 0)
                            {
                                rowXs.Add(htmlele);
                                prev_x = x;
                                prev_y = y;
                                prevele = htmlele;
                                prevHeight = prevele.ClientRectangle.Height;
                            }
                            else
                            {
                                if (colCount == 1)
                                {
                                    rowXs.Add(htmlele);
                                    diffX = x - (prev_x + prevele.ClientRectangle.Width);
                                    prev_x = x;
                                    prevHeight = (prevHeight + prevele.ClientRectangle.Height) / 2;
                                    prevele = htmlele;
                                }
                                else
                                {
                                    int _diffX = x - (prev_x + prevele.ClientRectangle.Width);
                                    if (Math.Abs(diffX - _diffX) <= 20)
                                    {
                                        rowXs.Add(htmlele);
                                        prev_x = x;
                                        prevHeight = (prevHeight + prevele.ClientRectangle.Height) / 2;
                                        prevele = htmlele;
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }

                            }
                            row.Add(htmlele);
                            colCount++;
                        }

                        if (colCount != 0)
                        {
                            if (rowGroup.Count == 0)
                            {
                                rowGroup.Add(row);
                                foreach (HtmlElement x in rowXs)
                                {
                                    matrix[y].Remove(x);
                                    matrixCount--;
                                }
                            }
                            else
                            {
                                if (rowGroup[0].Count == row.Count)
                                {
                                    rowGroup.Add(row);
                                    foreach (HtmlElement x in rowXs)
                                    {
                                        matrix[y].Remove(x);
                                        matrixCount--;
                                    }
                                    if (matrixCount == 0)
                                    {
                                        if (((rowGroup[0].Count * rowGroup.Count) > 4) && !((rowGroup.Count == 1) || (rowGroup[0].Count == 1)))
                                            rowGroups.Add(rowGroup);
                                    }
                                }
                                else
                                {
                                    if (((rowGroup[0].Count * rowGroup.Count) > 4) && !((rowGroup.Count == 1) || (rowGroup[0].Count == 1)))
                                    {
                                        rowGroups.Add(rowGroup);
                                    }
                                    else
                                    {
                                        rowGroup.Clear();
                                        rowGroup.Add(row);
                                        foreach (HtmlElement x in rowXs)
                                        {
                                            matrix[y].Remove(x);
                                            matrixCount--;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            foreach (List<List<HtmlElement>> htmleleGroup in rowGroups)
            {
                CreateButton_Level(htmleleGroup);
            }
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

            Plexiglass plexiglass = new Plexiglass(this, x - docLocation.X, y + 33 - docLocation.Y,
                /*x + htmlele.OffsetRectangle.Width > this.Width ? this.Width - x - 30 :*/ htmlele.OffsetRectangle.Width,
                /*y + htmlele.OffsetRectangle.Height > this.Height ? this.Height - y :*/ htmlele.OffsetRectangle.Height);
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
                /*x + parentX + htmlele_MSHTML.offsetWidth > this.Width ? this.Width - (x + parentX) - 30 :*/ htmlele_MSHTML.offsetWidth,
                /*y + parentY + htmlele_MSHTML.offsetHeight > this.Height ? this.Height - (y + parentY) :*/ htmlele_MSHTML.offsetHeight);
                //htmlele_MSHTML.offsetWidth, htmlele_MSHTML.offsetHeight);
            plexiglass._offsetLeft = x + parentX;
            plexiglass._offsetTop = y + parentY;
            plexiglassDictionary_MSHTML.Add(htmlele_MSHTML, plexiglass);
        }

        private void CreateButton_Level(List<List<HtmlElement>> htmleleGroup)
        {
            HtmlElement firstele = htmleleGroup[0][0], endele = htmleleGroup[htmleleGroup.Count - 1][htmleleGroup[0].Count - 1];
            int firstX, firstY, endX, endY, width, height;
            SetElementXY(firstele, out firstX, out firstY);
            SetElementXY(endele, out endX, out endY);
            width = endX + endele.ClientRectangle.Width - firstX;
            height = endY + endele.ClientRectangle.Height - firstY;
            if (((width > 0) && (height > 0)) && ((firstX >= 0) && (firstY >= 0)))
            {
                //TODO
                if (!plexiglassDictionary_Level.ContainsKey(htmleleGroup))
                {
                    Button newButton = new Button();
                    Point point = new Point(firstX < 20 ? 0 - docLocation.X : firstX - 20 - docLocation.X, firstY < 20 ? 0 - docLocation.Y : firstY - 20 - docLocation.Y);
                    newButton.Location = point;
                    newButton.Width = 15;
                    newButton.Height = 15;
                    newButton.Text = "V";
                    newButton.BackColor = Color.LightGray;

                    if ((docLocation.Y <= firstY) && (firstY + 20 <= webBrowser.Height))
                    {
                        newButton.Visible = true;
                    }
                    else
                    {
                        newButton.Visible = false;
                    }

                    newButton.Click += newButton_Click_Level;
                    newButton.MouseEnter += newButton_MouseEnter_Level;
                    newButton.MouseLeave += newButton_MouseLeave_Level;
                    this.Controls.Add(newButton);
                    newButton.Parent = webBrowser;

                    buttonTargetDictionary_Level.Add(newButton, htmleleGroup);

                    Plexiglass plexiglass = new Plexiglass(this, firstX - docLocation.X, firstY + 33 - docLocation.Y,
                        /*firstX + width > this.Width ? this.Width - firstX - 30 :*/ width,
                        /*firstY + height > this.Height ? this.Height - firstY :*/ height);
                    plexiglass._offsetLeft = firstX;
                    plexiglass._offsetTop = firstY;
                    plexiglassDictionary_Level.Add(htmleleGroup, plexiglass);
                }
            }
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

        void newButton_Click(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                newButton.BackColor = Color.Aqua;
                plexiglassDictionary[buttonTargetDictionary[newButton]].Visible = true;
                regionSelected++;
            }
            else
            {
                newButton.BackColor = Color.LightGray;
                plexiglassDictionary[buttonTargetDictionary[newButton]].Visible = false;
                regionSelected--;
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

        void newButton_Click_MSHTML(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                newButton.BackColor = Color.Aqua;
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].Visible = true;
                regionSelected++;
            }
            else
            {
                newButton.BackColor = Color.LightGray;
                plexiglassDictionary_MSHTML[buttonTargetDictionary_MSHTML[newButton]].Visible = false;
                regionSelected--;
            }
        }

        void newButton_MouseLeave_Level(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].BackColor = Color.LightSkyBlue;
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].Visible = false;
            }
            else
            {
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].BackColor = Color.LightSkyBlue;
            }
        }

        void newButton_MouseEnter_Level(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].BackColor = Color.LightBlue;
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].Visible = true;
            }
            else
            {
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].BackColor = Color.LightBlue;
            }
            this.ResumeLayout();
        }

        void newButton_Click_Level(object sender, EventArgs e)
        {
            Button newButton = sender as Button;
            if (newButton.BackColor.Equals(Color.LightGray))
            {
                newButton.BackColor = Color.Aqua;
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].Visible = true;
                regionSelected++;
            }
            else
            {
                newButton.BackColor = Color.LightGray;
                plexiglassDictionary_Level[buttonTargetDictionary_Level[newButton]].Visible = false;
                regionSelected--;
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
            alert_Init.Dispose();
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
    }
}