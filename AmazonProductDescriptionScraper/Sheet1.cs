using System;
using HtmlAgilityPack;

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

            int i = 1;


            while ( !(Cells[i, 1].value == null) && !(Cells[i, 1].value == "") )
            {
                string Url = Cells[i, 1].value;
                Globals.Sheet1.Cells[i, 2].value = GetProductDescription(Url);
                i++;
            }

        }

        private String GetProductDescription(String Url)
        {
            try
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument AmazonHtml = web.Load(Url);

                HtmlNode div = AmazonHtml.DocumentNode.SelectSingleNode(@"//div[@id=""productDescription""]");
               
                string Description = div.SelectSingleNode("//p").InnerText;

                return Description;

            }
            catch
            {
                return null;
            }
        }
    }
}
