using System;
using System.Linq;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using Remotion.Data.Linq.Clauses.ResultOperators;

namespace ReportDataCollector
{
    public class Client
    {
        public string Title { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
        public string CellNumber { get; set; }
        public string LandLine { get; set; }
        public DateTime DOB { get; set; }
        public string IDNumber { get; set; }
        public string HomeLanguage { get; set; }
        public string AssessmentLanguage { get; set; }
        public bool DriverLicense { get; set; }

        public string Gender
        {
            get { return GetGender(IDNumber); }
        }

        public string HeShe
        {
            get { return (Gender == "Female") ? "She" : "He"; }
        }

        private string GetGender(string idNumber)
        {
            if (string.IsNullOrEmpty(IDNumber)) throw new Exception("id number not specified");
            int lastDigit;
            if (!int.TryParse(idNumber.Last().ToString(), out lastDigit)) throw new Exception("id number not a valid format");
            return (lastDigit % 2 == 0) ? "Female" : "Male";
        }
    }


}
