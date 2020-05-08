using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using PdfReport.Reporting;
using PdfReport.Reporting.MigraDoc;

namespace PdfReport.Test
{
	internal class Program
	{
		private static int start = 0;
		private static int counter = 0;
		//const int noOfRecordsPerFile = 20000;
		const int noOfRecordsPerFile = 10;
		//static int TotalRecords = 100000;
		static int TotalRecords = 100;

		private static void Main()
		{
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			Stopwatch sw = new Stopwatch();
			var noOfIterations = TotalRecords / noOfRecordsPerFile;
			for (var i = 0; i < noOfIterations; i++)
			{
				var reportService = new ReportPdf();

				sw.Restart();
				var reportData = CreateReportData();
				var path = GetTempPdfPath();
				reportService.Export(path, reportData);
				sw.Stop();
				Console.WriteLine("Time taken" + sw.Elapsed.TotalSeconds);
				//Process.Start(path);
			}
			sw.Reset();
			sw.Start();
			Console.WriteLine();
			ReportPdf.CombinePDFs(@"D:\APDFData", "ConcatenatedDocument1_tempfile.pdf");
			sw.Stop();
			Console.WriteLine("Time taken to combine the pdf" + sw.Elapsed.TotalSeconds);
			Console.ReadLine();
		}

		private static ReportData CreateReportData()
		{
			return new ReportData
			{
				ClientReport = new ClientReport
				{

					ClientId = "1",
					ClientName = "Symphony Media -Linear",
					ReportParameters = "MVPD:All"

				},
				StructureSet = GetStructureData()
			};
		}

		private static StructureSet GetStructureData()
		{
			var billingReportObject = new StructureSet
			{
				ReportParameters = "MVPD: All",
				BillingReportStructure = new List<BillingReportStructure>
				{
						new BillingReportStructure
						{
							Invoice_Date = DateTime.Now.ToString("MM/yyyy"),
							Invoice_Number = 1,
							Agreement_Name = "Agreement1",
							MVPD="MVPD1",
							Service="Service1",
							SubscriberType="SubscriberType1",
							SystemDesignation="SystemDesignation1",
							Subs=100.00,
							Rate=0.5,
							Revenue=50.00,
						},
						new BillingReportStructure
						{
							Invoice_Date = DateTime.Now.ToString("MM/yyyy"),
							Invoice_Number = 1,
							Agreement_Name = "Agreement1",
							MVPD="MVPD1",
							Service="Service1",
							SubscriberType="SubscriberType1",
							SystemDesignation="SystemDesignation1",
							Subs=100.00,
							Rate=0.5,
							Revenue=50.00,
						},
						new BillingReportStructure
						{
							Invoice_Date = DateTime.Now.ToString("MM/yyyy"),
							Invoice_Number = 1,
							Agreement_Name = "Agreement1",
							MVPD="MVPD1",
							Service="Service1",
							SubscriberType="SubscriberType1",
							SystemDesignation="SystemDesignation1",
							Subs=100.00,
							Rate=0.5,
							Revenue=50.00,
						},
						new BillingReportStructure
						{
							Invoice_Date = DateTime.Now.ToString("MM/yyyy"),
							Invoice_Number = 1,
							Agreement_Name = "Agreement1",
							MVPD="MVPD1",
							Service="Service1",
							SubscriberType="SubscriberType1",
							SystemDesignation="SystemDesignation1",
							Subs=100.00,
							Rate=0.5,
							Revenue=50.00,
						},
						new BillingReportStructure
						{
							Invoice_Date = DateTime.Now.ToString("MM/yyyy"),
							Invoice_Number = 1,
							Agreement_Name = "Agreement2",
							MVPD="MVPD1",
							Service="Service1",
							SubscriberType="SubscriberType1",
							SystemDesignation="SystemDesignation1",
							Subs=100.00,
							Rate=0.5,
							Revenue=50.00,
						},
						new BillingReportStructure
						{
							Invoice_Date = DateTime.Now.ToString("MM/yyyy"),
							Invoice_Number = 1,
							Agreement_Name = "Agreement2",
							MVPD="MVPD1",
							Service="Service1",
							SubscriberType="SubscriberType1",
							SystemDesignation="SystemDesignation1",
							Subs=100.00,
							Rate=0.5,
							Revenue=50.00,
						}
					}
			};

			var arr = new BillingReportStructure[noOfRecordsPerFile];
			for (var i = start; i < noOfRecordsPerFile; i++)
			{
				arr[i] = new BillingReportStructure
				{
					Invoice_Date = DateTime.Now.ToString("MM/yyyy"),
					Invoice_Number = counter + 10,
					Agreement_Name = "Agreement" + counter.ToString(),
					MVPD = "MVPD" + counter.ToString(),
					Service = "Service" + counter.ToString(),
					SubscriberType = "SubscriberType" + counter.ToString(),
					SystemDesignation = "SystemDesignation" + counter.ToString(),
					Subs = 100.00 * counter,
					Rate = 0.5,
					Revenue = 100.00 * counter * 0.5,
				};
				counter++;
			}
			billingReportObject.BillingReportStructure = arr.ToList();
			return billingReportObject;
		}

		private static string GetTempPdfPath()
		{
			return Path.Combine(@"D:\APDFData", Guid.NewGuid().ToString() + ".pdf");
		}


	}
}
