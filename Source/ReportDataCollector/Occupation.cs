using System.Collections.Generic;

namespace ReportDataCollector
{
    public class Occupation
    {
        public string OccupationAtTimeOfIllness{ get; set; }
        public string EmployerAtTimeOfIllness{ get; set; }
        public string CurrentOccupation{ get; set; }
        public string CurrentEmployer{ get; set; }
        public string ClassificationOfWork { get; set; }

        public List<string> JobDescription { get; set; }
    }
}