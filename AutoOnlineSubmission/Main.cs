//#define HTMLPost
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Http;
using mshtml;

namespace AutoOnlineSubmission {
    public partial class Main : Form {
#if HTMLPost
        PostWebForm postWebForm = new PostWebForm();
#else
        private static readonly HttpClient client = new HttpClient();
#endif
        /// <summary>
        /// Webpage to enter and submit details contact details
        /// </summary>
        public const string WebPage = "https://agileautomations.co.uk/home/inputform";

        /// <summary>
        /// Contact details to be submitted
        /// </summary>
        public static Dictionary<string, string> ContactDetails = new Dictionary<string, string>();

        /// <summary>
        /// Used for extracting the reference digits from the response
        /// </summary>
        public static readonly int ReferenceStartIndex = "Reference - ".ToString().Length;

        /// <summary>
        /// Excel column index to save Reference # in
        /// </summary>
        private const int REFNO_COLUMN_IDX = 5;

        /// <summary>
        /// Names (id) of the fields on the webpage
        /// </summary>
        private class FormFieldNames {
            public const string NAME = "ContactName";
            public const string EMAIL = "ContactEmail";
            public const string SUBJECT = "ContactSubject";
            public const string MESSAGE = "Message";
        }

        public Main() {
            InitializeComponent();
        }

        private async void Main_Load(object sender, EventArgs e) {
            ofdExcelDoc.ShowDialog();
            if (ofdExcelDoc.CheckFileExists) {
                try {
                    //Create COM references
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ofdExcelDoc.FileName);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    //Form values for posting
                    //Dictionary<string, string> contactDetails = new Dictionary<string, string>();

                    int rowCount = xlRange.Rows.Count + 1;
                    int columnCount = xlRange.Columns.Count + 1;
                    for (int row = 1; row < rowCount; row++) {
                        for (int col = 1; col < columnCount; col++) {
                            //Build form values from excel document
                            if (xlRange.Cells[row, col] != null && xlRange.Cells[row, col].Value2 != null) {
                                switch (col) {
                                    case 1: ContactDetails.Add(FormFieldNames.NAME, xlRange.Cells[row, 1].value2.ToString()); break;
                                    case 2: ContactDetails.Add(FormFieldNames.EMAIL, xlRange.Cells[row, 2].value2.ToString()); break;
                                    case 3: ContactDetails.Add(FormFieldNames.SUBJECT, xlRange.Cells[row, 3].value2.ToString()); break;
                                    case 4: ContactDetails.Add(FormFieldNames.MESSAGE, xlRange.Cells[row, 4].value2.ToString()); break;
                                    default: col = columnCount; break;  //Done building collection
                                }
                            }
                        }

#if HTMLPost
                        //Post current contact details to the webpage
                        postWebForm.ShowDialog();

                        //Store the reference number in the excel sheet for this contact
                        if (PostWebForm.HtmlRefno.Length > 0) {
                            xlWorksheet.Cells[row, REFNO_COLUMN_IDX] = PostWebForm.HtmlRefno;
                            xlWorkbook.Save();
                        }
#else
                        //Post contact details and store HTML response
                        string response = await SubmitForm(ContactDetails);

                        //Anything in the response is worth saving
                        if (response.Length > 0) {
                            //Store the reference number in the excel sheet for this contact
                            xlWorksheet.Cells[row, REFNO_COLUMN_IDX] = ExtractReferenceNo(response);
                            xlWorkbook.Save();
                        }
#endif
                        ContactDetails.Clear();
                    }

                    //Display the Excel document
                    xlApp.Visible = true;

                    //Release COM Object
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.ReleaseComObject(xlRange);
                    Marshal.FinalReleaseComObject(xlWorksheet);
                    Marshal.ReleaseComObject(xlWorkbook);
                    Marshal.ReleaseComObject(xlApp);
                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
            }
            this.Close();
        }

#if !HTMLPost
        /// <summary>
        /// Post contact details
        /// </summary>
        /// <param name="_values">Contact form details</param>
        /// <returns>HTML response if successful, otherwise empty string</returns>
        private async Task<string> SubmitForm(Dictionary<string, string> _values) {
            FormUrlEncodedContent content = new FormUrlEncodedContent(_values);
            HttpResponseMessage response = await client.PostAsync(WebPage, content);
            if (response.IsSuccessStatusCode) return await response.Content.ReadAsStringAsync();
            else return "";
        }

        /// <summary>
        /// Extract the refno from Label Id ReferenceNo
        /// </summary>
        /// <param name="_html">HTML response</param>
        /// <returns>Reference number</returns>
        private string ExtractReferenceNo(string _html) {
            HTMLDocument doc = new HTMLDocument();
            IHTMLDocument2 doc2 = (IHTMLDocument2)doc;
            doc2.write(_html);      //Load string as HTML document
            string innerText = doc.getElementById("ReferenceNo").innerText;
            return innerText.Substring(ReferenceStartIndex, innerText.Length - ReferenceStartIndex);
        }
#endif
    }
}
