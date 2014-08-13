using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using NGS.Templater;
using OTS.Domain;

namespace ReportGeneration
{
    public class AsposeWordGenerator
    {
        private readonly string _rootPath;

        public AsposeWordGenerator(string rootPath)
        {
            _rootPath = rootPath;
        }

        public void GenerateReport(ReportData reportData)
        {
            var path = @"c:\googledrive\report.docx";
            Aspose.Words.Document template = new Document(Path.Combine(_rootPath, reportData.TemplateName + ".odt"));
            foreach (string snippet in reportData.Snippets.Where(snippet => !string.IsNullOrEmpty(snippet)))
            {
                Document snippetDocument = new Document(GetPathForSnippet(snippet,reportData.TemplateName,reportData.Therapist));
                template.AppendDocument(snippetDocument,ImportFormatMode.KeepSourceFormatting);
             }
            template.Save(path);

            using (ITemplateDocument templateDocument = Configuration.Factory.Open(path))
            {
                templateDocument.Process(reportData);

            }
            // remove license issue
            var document = new Document(path);
            document.Range.Replace("Unlicensed version. Please register @ templater.info", string.Empty, true, false);
            document.Save(path);

        }

        public string GetPathForSnippet(string bookmark, string templateName, string therapist)
        {
            string ext = ".odt";
            var snippetPath = Path.Combine(_rootPath, "snippets");
            string bookmark_Template_Therapist_Path = Path.Combine(snippetPath,String.Format("{0}_{1}_{2}{3}",bookmark,templateName ,therapist,ext));
            string bookmark_Therapist_Path = Path.Combine(snippetPath,String.Format("{0}_{1}{2}", bookmark, therapist,ext));
            string bookmark_Template_Path = Path.Combine(snippetPath,String.Format("{0}_{1}{2}", bookmark, templateName,ext));
            string bookmark_Path = Path.Combine(snippetPath, bookmark + ext);

            var bookList = new List<string>() { bookmark_Template_Therapist_Path, bookmark_Therapist_Path, bookmark_Template_Path,bookmark_Path};
            return bookList.Where(File.Exists).First();
        }
    }
}
