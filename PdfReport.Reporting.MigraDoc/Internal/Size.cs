using MigraDoc.DocumentObjectModel;

namespace PdfReport.Reporting.MigraDoc.Internal
{
    internal class Size
    {
        // Top and bottom margins are larger to account for the header and footer

        public static readonly Unit TopPageMargin = "2.5 in";
        public static readonly Unit BottomPageMargin = "0.5 in";
        public static readonly Unit LeftRightPageMargin = "0.50 in";

        public static readonly Unit HeaderMargin = "0 in";
        public static readonly Unit FooterMargin = "0 in";

        public static readonly Unit TableCellPadding = "0.07 in";

        public static Unit GetWidth(Section section)
        {
            section.PageSetup.PageFormat = PageFormat.P11x17;
            PageSetup.GetPageSize(section.PageSetup.PageFormat, out Unit pageWidth, out Unit _);
            return pageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
        }
    }
}