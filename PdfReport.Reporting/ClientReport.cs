using System;

namespace PdfReport.Reporting
{
    public class ClientReport
    {
        //public string Id { get; set; }
        //public string FirstName { get; set; }
        //public string LastName { get; set; }
        //public Sex Sex { get; set; }
        //public DateTime Birthdate { get; set; }
        //public Doctor Doctor { get; set; }

        public string ClientId { get; set; }
        public string ClientName { get; set; }

        public string Logo { get; set; }

        public string ReportTitle { get; set; }
        public string ReportParameters { get; set; }
    }
}