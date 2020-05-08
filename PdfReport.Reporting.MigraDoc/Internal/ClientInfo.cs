using System;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace PdfReport.Reporting.MigraDoc.Internal
{
    internal class ClientInfo
    {
        public static readonly Color Shading = new Color(243, 243, 243);

        public void Add(Section section, ClientReport clientReport)
        {
            //var table = AddPatientInfoTable(section);

            //AddLeftInfo(table.Rows[0].Cells[0], clientReport);
            //AddLeftInfo(table.Rows[0], clientReport);
            //AddRightInfo(table.Rows[0].Cells[1], clientReport);
        }

        private Table AddPatientInfoTable(Section section)
        {
            var table = section.AddTable();
            table.Shading.Color = Shading;

            table.Rows.LeftIndent = 0;

            table.LeftPadding = Size.TableCellPadding;
            table.TopPadding = Size.TableCellPadding;
            table.RightPadding = Size.TableCellPadding;
            table.BottomPadding = Size.TableCellPadding;

            // Use two columns of equal width
            var columnWidth = Size.GetWidth(section) / 2.0;
            table.AddColumn(columnWidth);
            table.AddColumn(columnWidth);

            // Only one row is needed
            table.AddRow();

            return table;
        }

        
       

        private void AddRightInfo(Cell cell, ClientReport clientReport)
        {
            var p = cell.AddParagraph();

            // Add birthdate
            p.AddText("Client Name: ");
            p.AddFormattedText((clientReport.ClientName), TextFormat.Bold);

            p.AddLineBreak();

            // Add doctor name
            p.AddText(" ");
            p.AddFormattedText($"{clientReport.ReportParameters}", TextFormat.Bold);
        }

        private string Format(DateTime birthdate)
        {
            return $"{birthdate:d} (age {Age(birthdate)})";
        }

        // See http://stackoverflow.com/a/1404/1383366
        private int Age(DateTime birthdate)
        {
            var today = DateTime.Today;
            int age = today.Year - birthdate.Year;
            return birthdate.AddYears(age) > today ? age - 1 : age;
        }
    }
}