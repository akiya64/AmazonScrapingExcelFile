using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Sgml;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.Net;

namespace AmazonProductDescriptionScraper
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {

            Range Current = Globals.Sheet1.Range["A1:A100"];

            foreach (Range r in Current)
            {
                Globals.Sheet1.Cells[r.Row, 2].value = GetProductDescription(r.Value);
            }

        }

        private String GetProductDescription(String Url)
        {

            Sgml.SgmlReader sgmlReader = new Sgml.SgmlReader();
            sgmlReader.DocType = "HTML";
            sgmlReader.WhitespaceHandling = WhitespaceHandling.All;
            sgmlReader.CaseFolding = Sgml.CaseFolding.ToLower;

            sgmlReader.InputStream = FetchHtmlDoc(Url);

            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.XmlResolver = null;
            doc.Load(sgmlReader);

            XmlNodeList Pnodes = doc.GetElementsByTagName("p");

            String Description = Pnodes[0].InnerText;

            return Description;

        }

        private TextReader FetchHtmlDoc(String Url)
        {
            WebRequest request = HttpWebRequest.Create(Url);
            WebResponse response = request.GetResponse();

            TextReader Html = new StreamReader(response.GetResponseStream());


            return Html ;
        }
    }
}
