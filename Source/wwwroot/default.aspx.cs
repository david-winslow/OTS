using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ReportGeneration;

namespace wwwroot
{
    public partial class _default : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack)
            {
                var excelFile = this.excelFile.PostedFile;
                var templateFile = this.wordTemplateFile.PostedFile;

                var excel = Path.Combine(Server.MapPath("~/Content/report"), excelFile.FileName);

                String template = string.Empty;
                if (templateFile != null && templateFile.ContentLength > 0)
                {
                    template = Path.Combine(Server.MapPath("~/Content/report"), templateFile.FileName);
                    templateFile.SaveAs(template);
                }
                else
                {
                    template = Path.Combine(Server.MapPath("~/Content/report"), "template.docx");
                }
                string fileName = string.Format("report{0}.docx", DateTime.Now.Second);
                string destFileName = Path.Combine(Server.MapPath("~/Content/report"), fileName);
                excelFile.SaveAs(excel);

                File.Copy(template, destFileName, true);
                WordGenerator gen = new WordGenerator(destFileName, excel);
                gen.Generate();
                HyperLink1.NavigateUrl = string.Format("~/content/report/{0}", fileName);
                HyperLink1.Visible = true;

            }
        }
    }
}