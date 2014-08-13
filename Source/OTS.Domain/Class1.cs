using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OTS.Domain
{
    public class ReportData
    {
        public DateTime Date { get; set; }
        public string Surname { get; set; }
        public string FirstName { get; set; }
        public string Therapist { get; set; }
        public string TemplateName { get; set; }
        public IEnumerable<string> Snippets { get; set; }
    }
}
