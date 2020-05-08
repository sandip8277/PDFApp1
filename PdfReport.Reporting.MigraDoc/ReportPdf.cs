﻿using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfReport.Reporting.MigraDoc.Internal;
using System.IO;

namespace PdfReport.Reporting.MigraDoc
{
    public class ReportPdf : IReport
    {
        public void Export(string path, ReportData data)
        {
            ExportPdf(path, CreateReport(data));
        }

        private void ExportPdf(string path, Document report)
        {
            var pdfRenderer = new PdfDocumentRenderer();
            pdfRenderer.Document = report;
            pdfRenderer.RenderDocument();
            pdfRenderer.PdfDocument.Save(path);
        }

        private Document CreateReport(ReportData data)
        {
            var doc = new Document();
            CustomStyles.Define(doc);
            doc.Add(CreateMainSection(data));
            return doc;
        }

        private Section CreateMainSection(ReportData data)
        {
            var section = new Section();
            SetUpPage(section);
            AddHeaderAndFooter(section, data);
            AddContents(section, data);
            return section;
        }

        private void SetUpPage(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.Letter;
            section.PageSetup.Orientation = Orientation.Landscape;
            section.PageSetup.LeftMargin = Size.LeftRightPageMargin;
            section.PageSetup.RightMargin = Size.LeftRightPageMargin;
            section.PageSetup.TopMargin = Size.TopPageMargin;
            section.PageSetup.BottomMargin = Size.BottomPageMargin;
            section.PageSetup.HeaderDistance = Size.HeaderMargin;
            section.PageSetup.FooterDistance = Size.FooterMargin;
        }

        private void AddHeaderAndFooter(Section section, ReportData data)
        {
            new HeaderAndFooter().Add(section, data);
        }

        private void AddContents(Section section, ReportData data)
        {
            AddPatientInfo(section, data.ClientReport);
            AddStructureSet(section, data.StructureSet);
        }

        private void AddPatientInfo(Section section, ClientReport clientReport)
        {
            new ClientInfo().Add(section, clientReport);
        }

        private void AddStructureSet(Section section, StructureSet structureSet)
        {
            new StructureSetContent().Add(section, structureSet);
        }

        public static void CombinePDFs(string path, string outputFileName)
        {
            // Get some file names
            string[] files = System.IO.Directory.GetFiles(path, "*.pdf");

            // Open the output document
            PdfDocument outputDocument = new PdfDocument();

            // Iterate files
            foreach (string file in files)
            {
                // Open the document to import pages from it.
                PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import);

                // Iterate pages
                int count = inputDocument.PageCount;
                for (int idx = 0; idx < count; idx++)
                {
                    // Get the page from the external document...
                    PdfPage page = inputDocument.Pages[idx];
                    // ...and add it to the output document.
                    outputDocument.AddPage(page);
                }
            }

            // Save the document...
            string filename = Path.Combine(path, outputFileName);
            outputDocument.Save(filename);
            // ...and start a viewer.

        }
    }
}
