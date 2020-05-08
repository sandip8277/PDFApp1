using System.Collections.Generic;

namespace PdfReport.Reporting
{
    public class StructureSet
    {
        public string ReportParameters { get; set; }
        public List<BillingReportStructure> BillingReportStructure { get; set; }
    }
}