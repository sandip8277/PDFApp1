using MigraDoc.DocumentObjectModel;

namespace PdfReport.Reporting.MigraDoc.Internal
{
    internal class CustomStyles
    {
        public const string ReportName = "ReportName";
        public const string ColumnHeader = "ColumnHeader";

        public static void Define(Document doc)
        {
            var reportName = doc.Styles.AddStyle(ReportName, StyleNames.Normal);
            reportName.ParagraphFormat.Font.Size = 10;
            reportName.ParagraphFormat.Font.Bold = true;

            var heading1 = doc.Styles[StyleNames.Heading4];
            heading1.BaseStyle = StyleNames.Normal;
            heading1.Font.Size = 10;
            heading1.ParagraphFormat.SpaceBefore = 20;

            var heading2 = doc.Styles[StyleNames.Heading2];
            heading2.BaseStyle = StyleNames.Normal;
            heading2.ParagraphFormat.Shading.Color = Color.FromRgb(0, 0, 0);
            heading2.ParagraphFormat.Font.Color = Color.FromRgb(255, 255, 255);
            heading2.ParagraphFormat.Font.Bold = true;
            heading2.ParagraphFormat.SpaceBefore = 10;
            heading2.ParagraphFormat.LeftIndent = Size.TableCellPadding;
            heading2.ParagraphFormat.RightIndent = Size.TableCellPadding;
            heading2.ParagraphFormat.Borders.Distance = Size.TableCellPadding;

            var columnHeader = doc.Styles.AddStyle(ColumnHeader, StyleNames.Normal);
            columnHeader.ParagraphFormat.Font.Bold = true;
            columnHeader.ParagraphFormat.LeftIndent = Size.TableCellPadding;
            columnHeader.ParagraphFormat.RightIndent = Size.TableCellPadding;
        }
    }
}