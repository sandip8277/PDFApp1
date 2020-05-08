using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System.Collections.Generic;
using System.Linq;


namespace PdfReport.Reporting.MigraDoc.Internal
{
    internal class StructureSetContent
    {
        public void Add(Section section, StructureSet structureSet)
        {
            //AddHeading(section, structureSet);
            AddStructures(section, structureSet.BillingReportStructure);
        }

        //    private void AddHeading(Section section, StructureSet structureSet)
        //    {
        //  section.AddParagraph(structureSet.ReportParameters, StyleNames.Heading1);
        //  //section.AddParagraph($"Image {structureSet.Image.Id} " +
        //  //                     $"taken {structureSet.Image.CreationTime:g}");
        //  section.AddParagraph("Report Date Range: 04/2019-05/2020");
        //  section.AddParagraph(structureSet.ReportParameters);
        //  section.AddParagraph("Service:All");
        //  section.AddParagraph("----------------------------------------------------------------------------------");

        //}

        private void AddStructures(Section section, List<BillingReportStructure> structures)
        {
            //AddTableTitle(section, "(SymphonyMedia-Linear) Billing Summary by Agreement");
            AddStructureTable(section, structures);
        }

        private void AddTableTitle(Section section, string title)
        {
            var p = section.AddParagraph(title, StyleNames.Heading2);
            p.Format.KeepWithNext = true;
        }

        private void AddStructureTable(Section section, List<BillingReportStructure> structures)
        {
            var table = section.AddTable();

            FormatTable(table);
            AddColumnsAndHeaders(table);
            AddStructureRows(table, structures);

            AddLastRowBorder(table);
            AlternateRowShading(table);
        }

        private static void FormatTable(Table table)
        {
            table.LeftPadding = 0;
            table.TopPadding = Size.TableCellPadding;
            table.RightPadding = 0;
            table.BottomPadding = Size.TableCellPadding;
            table.Format.LeftIndent = Size.TableCellPadding;
            table.Format.RightIndent = Size.TableCellPadding;
        }

        private void AddColumnsAndHeaders(Table table)
        {
            var width = Size.GetWidth(table.Section);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);
            table.AddColumn(width * 0.16);

            var headerRow = table.AddRow();
            headerRow.Borders.Bottom.Width = 1;

        }

        private void AddHeader(Cell cell, string header)
        {
            var p = cell.AddParagraph(header);
            p.Style = CustomStyles.ColumnHeader;
            cell.Shading.Color = Color.FromRgb(161, 161, 161);
        }

        private void AddStructureRows(Table table, List<BillingReportStructure> structures)
        {
            List<BillingReportStructure> st = structures;
            st.OrderBy(x => x.Agreement_Name);
            IEnumerable<string> lstDistictsAgreement;
            string previousAgreement = "";
            int nbrOfRecord = 0,
              temp = 0;
            double grandSubTotal = 0;
            double grandRateTotal = 0;
            double grandRevenueTotal = 0;
            lstDistictsAgreement = st.Select(x => x.Agreement_Name).Distinct();

            foreach (var item in lstDistictsAgreement)
            {
                nbrOfRecord = 0;
                temp++;
                for (int j = 0; j < st.Count; j++)
                {
                    if (st[j].Agreement_Name == item)
                    {
                        nbrOfRecord++;

                        if (j == 0 || st[j - 1].Agreement_Name != previousAgreement)
                        {
                            
                            var headerRow = table.AddRow();
                            headerRow.Borders.Bottom.Width = 1;
                            AddHeader(headerRow.Cells[0], "Invoice Date");
                            AddHeader(headerRow.Cells[1], "Invoice Number");
                            AddHeader(headerRow.Cells[2], "Agreement Name");
                            AddHeader(headerRow.Cells[3], "MVPD");
                            AddHeader(headerRow.Cells[4], "Service");
                            AddHeader(headerRow.Cells[5], "Subscriber Type");
                            AddHeader(headerRow.Cells[6], "System Designation");
                            AddHeader(headerRow.Cells[7], "Subs");
                            AddHeader(headerRow.Cells[8], "Rate");
                            AddHeader(headerRow.Cells[9], "Revenue");
                        }
                        if (j == 0)
                        {
                            if (nbrOfRecord == 1)
                            {
                                var row1 = table.AddRow();
                                row1.VerticalAlignment = VerticalAlignment.Center;

                                row1.Cells[0].AddParagraph(st[j].Invoice_Date);
                                row1.Cells[1].AddParagraph($"{st[j].Invoice_Number}");
                                row1.Cells[2].AddParagraph($"{st[j].Agreement_Name}");
                                row1.Cells[3].AddParagraph($"{st[j].MVPD}");
                                row1.Cells[4].AddParagraph($"{st[j].Service}");
                                row1.Cells[5].AddParagraph($"{st[j].SubscriberType:f2}");
                                row1.Cells[6].AddParagraph($"{st[j].SystemDesignation:f2}");
                                row1.Cells[7].AddParagraph($"{st[j].Subs:f2}");
                                row1.Cells[8].AddParagraph("$" + $"{st[j].Rate:f2}");
                                row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                            }
                            else
                            {
                              
                                var row1 = table.AddRow();
                                row1.VerticalAlignment = VerticalAlignment.Center;
                                row1.Cells[0].AddParagraph("");
                                row1.Cells[1].AddParagraph("");
                                row1.Cells[2].AddParagraph("");
                                row1.Cells[3].AddParagraph("");
                                row1.Cells[4].AddParagraph("");
                                row1.Cells[5].AddParagraph("");
                                row1.Cells[6].AddParagraph("");
                                row1.Cells[7].AddParagraph($"{st[j].Subs:f2}");
                                row1.Cells[8].AddParagraph("$" + $"{st[j].Rate:f2}");
                                row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                            }

                        }
                        else
                        {
                            if (st[j - 1].Agreement_Name != previousAgreement)
                            {
                                if (nbrOfRecord == 1)
                                {
                                    var row1 = table.AddRow();
                                    row1.VerticalAlignment = VerticalAlignment.Center;

                                    row1.Cells[0].AddParagraph(st[j].Invoice_Date);
                                    row1.Cells[1].AddParagraph($"{st[j].Invoice_Number}");
                                    row1.Cells[2].AddParagraph($"{st[j].Agreement_Name}");
                                    row1.Cells[3].AddParagraph($"{st[j].MVPD}");
                                    row1.Cells[4].AddParagraph($"{st[j].Service}");
                                    row1.Cells[5].AddParagraph($"{st[j].SubscriberType:f2}");
                                    row1.Cells[6].AddParagraph($"{st[j].SystemDesignation:f2}");
                                    row1.Cells[7].AddParagraph($"{st[j].Subs:f2}");
                                    row1.Cells[8].AddParagraph("$" + $"{st[j].Rate:f2}");
                                    row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                                }
                                else
                                {
                                    var row1 = table.AddRow();
                                    row1.VerticalAlignment = VerticalAlignment.Center;
                                    row1.Cells[0].AddParagraph("");
                                    row1.Cells[1].AddParagraph("");
                                    row1.Cells[2].AddParagraph("");
                                    row1.Cells[3].AddParagraph("");
                                    row1.Cells[4].AddParagraph("");
                                    row1.Cells[5].AddParagraph("");
                                    row1.Cells[6].AddParagraph("");
                                    row1.Cells[7].AddParagraph($"{st[j].Subs:f2}");
                                    row1.Cells[8].AddParagraph("$" + $"{st[j].Rate:f2}");
                                    row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                                }

                            }
                            else
                            {
                                if (nbrOfRecord == 1)
                                {
                                    var row1 = table.AddRow();
                                    row1.VerticalAlignment = VerticalAlignment.Center;

                                    row1.Cells[0].AddParagraph(st[j].Invoice_Date);
                                    row1.Cells[1].AddParagraph($"{st[j].Invoice_Number}");
                                    row1.Cells[2].AddParagraph($"{st[j].Agreement_Name}");
                                    row1.Cells[3].AddParagraph($"{st[j].MVPD}");
                                    row1.Cells[4].AddParagraph($"{st[j].Service}");
                                    row1.Cells[5].AddParagraph($"{st[j].SubscriberType:f2}");
                                    row1.Cells[6].AddParagraph($"{st[j].SystemDesignation:f2}");
                                    row1.Cells[7].AddParagraph($"{st[j].Subs:f2}");
                                    row1.Cells[8].AddParagraph("$" + $"{st[j].Rate:f2}");
                                    row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                                }
                                else
                                {
                                    var row1 = table.AddRow();
                                    row1.VerticalAlignment = VerticalAlignment.Center;
                                    row1.Cells[0].AddParagraph("");
                                    row1.Cells[1].AddParagraph("");
                                    row1.Cells[2].AddParagraph("");
                                    row1.Cells[3].AddParagraph("");
                                    row1.Cells[4].AddParagraph("");
                                    row1.Cells[5].AddParagraph("");
                                    row1.Cells[6].AddParagraph("");
                                    row1.Cells[7].AddParagraph($"{st[j].Subs:f2}");
                                    row1.Cells[8].AddParagraph("$" + $"{st[j].Rate:f2}");
                                    row1.Cells[9].AddParagraph("$" + $"{st[j].Revenue:f2}");
                                }
                            }
                        }

                    }
                    previousAgreement = item;
                    if (j == (st.Count() - 1))
                    {
                        var row1 = table.AddRow();
                        row1.VerticalAlignment = VerticalAlignment.Center;
                        row1.Borders.Top.Width = 2;
                        row1.Cells[0].AddParagraph("");
                        row1.Cells[1].AddParagraph("");
                        row1.Cells[2].AddParagraph("");
                        row1.Cells[3].AddParagraph("");
                        row1.Cells[4].AddParagraph("");
                        row1.Cells[5].AddParagraph("");
                        row1.Cells[6].AddParagraph($"Sub Total");
                        row1.Cells[7].AddParagraph($"{CalculateSubs(item, st, "Subs"):f2}");
                        row1.Cells[9].AddParagraph("$" + $"{CalculateSubs(item, st, "Revenue"):f2}");
                    }
                    if (j == (st.Count() - 1) && temp == lstDistictsAgreement.Count())
                    {
                        var row1 = table.AddRow();
                        row1.VerticalAlignment = VerticalAlignment.Center;
                        row1.Borders.Top.Width = 2;
                        row1.Format.Font.Bold = true;
                        row1.Format.Font.Size = 10;
                        row1.Cells[0].AddParagraph("");
                        row1.Cells[1].AddParagraph("");
                        row1.Cells[2].AddParagraph("");
                        row1.Cells[3].AddParagraph("");
                        row1.Cells[4].AddParagraph("");
                        row1.Cells[5].AddParagraph("");
                        row1.Cells[6].AddParagraph($"Grand Total");
                        row1.Cells[7].AddParagraph($"{CalculateTotalSubs(item, st, "Subs"):f2}");
                        //row1.Cells[8].AddParagraph("$" + CalculateSubs(item, st, "Rate"));
                        row1.Cells[9].AddParagraph("$" + $"{CalculateTotalSubs(item, st, "Revenue"):f2}");
                    }
                }

            }
        }
        private double CalculateSubs(string currentItem, List<BillingReportStructure> currentStruct, string columnName)
        {
            double totalSubs = 0;
            switch (columnName)
            {
                case "Subs":
                    {
                        totalSubs = currentStruct.Where(obj => obj.Agreement_Name == currentItem).Sum(obj => obj.Subs);
                        totalSubs.ToString(":f2");
                        break;
                    }
                case "Rate":
                    {
                        totalSubs = currentStruct.Where(obj => obj.Agreement_Name == currentItem).Sum(obj => obj.Rate);
                        break;
                    }
                case "Revenue":
                    {
                        totalSubs = currentStruct.Where(obj => obj.Agreement_Name == currentItem).Sum(obj => obj.Revenue);
                        break;
                    }

            }
            return totalSubs;
        }
        private double CalculateTotalSubs(string currentItem, List<BillingReportStructure> currentStruct, string columnName)
        {
            double grandTotalSubs = 0;
            switch (columnName)
            {
                case "Subs":
                    {
                        grandTotalSubs = currentStruct.Sum(obj => obj.Subs);
                        break;
                    }
                case "Rate":
                    {
                        grandTotalSubs = currentStruct.Sum(obj => obj.Rate);
                        break;
                    }
                case "Revenue":
                    {
                        grandTotalSubs = currentStruct.Sum(obj => obj.Revenue);
                        break;
                    }

            }
            return grandTotalSubs;
        }

        private void AddLastRowBorder(Table table)
        {
            var lastRow = table.Rows[table.Rows.Count - 1];
            lastRow.Borders.Bottom.Width = 2;
        }

        private void AlternateRowShading(Table table)
        {
            // Start at i = 1 to skip column headers
            for (var i = 1; i < table.Rows.Count; i++)
            {
                //if (i % 2 == 0)  // Even rows
                //{
                //    table.Rows[i].Shading.Color = Color.FromRgb(216, 216, 216);
                //}
                //else if (i == 1)
                //{
                //  table.Rows[i].Shading.Color = Color.FromRgb(216, 216, 216);

                //}
            }
        }

    }
}