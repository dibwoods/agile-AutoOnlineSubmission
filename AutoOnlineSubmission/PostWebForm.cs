using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoOnlineSubmission {
    public partial class PostWebForm : Form {
        /// <summary>
        /// Have the contact details been submitted?
        /// </summary>
        private bool submitted;

        /// <summary>
        /// Reference Number from submitted contact details
        /// </summary>
        public static string HtmlRefno;

        public PostWebForm() {
            InitializeComponent();
            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);
        }

        private void PostWebForm_Load(object sender, EventArgs e) {
            submitted = false;
            HtmlRefno = "";
            webBrowser1.Navigate(Main.WebPage);
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e) {
            if ((sender as WebBrowser).ReadyState == System.Windows.Forms.WebBrowserReadyState.Complete) {
                try {
                    if (!submitted) {
                        //Populate the contact details
                        foreach (KeyValuePair<string, string> kvp in Main.ContactDetails) {
                            webBrowser1.Document.GetElementById(kvp.Key).InnerText = kvp.Value;
                        }

                        //Click the submit button
                        var elements = webBrowser1.Document.GetElementsByTagName("input");
                        foreach (HtmlElement element in elements) {
                            if (element.OuterHtml.Contains("Send message")) {
                                element.InvokeMember("click");
                                break;
                            }
                        }
                        submitted = true;
                    }
                    else {
                        //After submitting contact details, grab the reference number
                        string refNo = webBrowser1.Document.GetElementById("ReferenceNo").InnerText;
                        HtmlRefno = refNo.Substring(Main.ReferenceStartIndex, refNo.Length - Main.ReferenceStartIndex);
                        this.Close();
                    }
                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
