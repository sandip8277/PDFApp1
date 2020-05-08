using System;

namespace PdfReport.Reporting
{
    public class BillingReportStructure
    {
       
        public string Invoice_Date { get; set; }
        public int Invoice_Number { get; set; }
        public string Agreement_Name { get; set; }

        public string MVPD { get; set; }
        public string Service { get; set; }
        public string SubscriberType { get; set; }
        public string SystemDesignation { get; set; }
        public double  Subs { get; set; }
        public double Rate { get; set; }
        public double Revenue { get; set; }

    }
}