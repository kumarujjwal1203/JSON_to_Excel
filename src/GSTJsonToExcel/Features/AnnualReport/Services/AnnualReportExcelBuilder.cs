using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using GSTJsonToExcel.Features.AnnualReport.Models;
using GSTJsonToExcel.Helpers;

namespace GSTJsonToExcel.Features.AnnualReport.Services
{
    public class AnnualReportExcelBuilder
    {
        private const string AccountingNumberFormat = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * \"-\"??_ ;_ @_";
        private static readonly XLColor HeaderBlueFill = XLColor.FromHtml("#0070C0");
        private static readonly XLColor SectionSummaryBlueFill = XLColor.FromHtml("#BDD7EE");
        private static readonly XLColor DifferenceGrayFill = XLColor.FromHtml("#D7D7D7");
        private static readonly XLColor LightBorderColor = XLColor.FromHtml("#D9D9D9");

        public string BuildWorkbook(
            AnnualWorkbookData data,
            AnnualReportUserSettings settings,
            string outputDirectory,
            List<string> includedSheets,
            List<string> skippedReasons,
            string? singleGstinFilter = null)
        {
            Directory.CreateDirectory(outputDirectory);

            using var wb = new XLWorkbook();
            wb.Style.Font.FontName = "Calibri";
            wb.Style.Font.FontSize = 9;

            // 12-month labels: Apr 2025 .. Mar 2026
            string[] labels12 = Get12MonthLabels(data.StartYear, data.EndYear);

            // 24-month labels: Apr 2025 .. Mar 2026 + Apr 2026 .. Mar 2027
            string[] labels24 = Get24MonthLabels(data.StartYear, data.EndYear, data.NextEndYear);

            Func<AnnualSheetRow, bool> rowFilter = r =>
                string.IsNullOrEmpty(singleGstinFilter) ||
                string.Equals(r.Gstin, singleGstinFilter, StringComparison.OrdinalIgnoreCase);

            // 1. Overview (Exact match to screenshot + Executive Summary Table below)
            if (settings.IncludeOverview)
            {
                BuildOverviewSheet(wb, data, singleGstinFilter);
                includedSheets.Add("Overview");
            }

            // 2. GSTR-3B (12 Months: Company GSTIN | Section | Type | Apr 2025 .. Mar 2026 | Total)
            if (settings.IncludeGstr3B)
            {
                var rows = data.Gstr3BRows.Where(rowFilter).ToList();
                if (data.Has3BData && rows.Count > 0)
                {
                    BuildMonthGridSheet(wb, "GSTR-3B", labels12, rows);
                    includedSheets.Add("GSTR-3B");
                }
                else
                {
                    skippedReasons.Add("GSTR-3B skipped (no GSTR-3B files in selection)");
                }
            }

            // 3. GSTR-1 (24 Months: Company GSTIN | Section | Type | Apr 2025 .. Mar 2026 | Apr 2026 .. Mar 2027 | Total)
            if (settings.IncludeGstr1)
            {
                var rows = data.Gstr1Rows.Where(rowFilter).ToList();
                if (data.Has1Data && rows.Count > 0)
                {
                    BuildMonthGridSheet(wb, "GSTR-1", labels24, rows);
                    includedSheets.Add("GSTR-1");
                }
                else
                {
                    skippedReasons.Add("GSTR-1 skipped (no GSTR-1 files in selection)");
                }
            }

            // 4. GSTR-2A (24 Months: Company GSTIN | Section | Type | Apr 2025 .. Mar 2026 | Apr 2026 .. Mar 2027 | Total)
            if (settings.IncludeGstr2A)
            {
                var rows = data.Gstr2ARows.Where(rowFilter).ToList();
                if (data.Has2AData && rows.Count > 0)
                {
                    BuildMonthGridSheet(wb, "GSTR-2A", labels24, rows);
                    includedSheets.Add("GSTR-2A");
                }
                else
                {
                    skippedReasons.Add("GSTR-2A skipped (no GSTR-2A files in selection)");
                }
            }

            // 5. GSTR-2B (24 Months: Company GSTIN | Section | Type | Apr 2025 .. Mar 2026 | Apr 2026 .. Mar 2027 | Total)
            if (settings.IncludeGstr2B)
            {
                var rows = data.Gstr2BRows.Where(rowFilter).ToList();
                if (data.Has2BData && rows.Count > 0)
                {
                    BuildMonthGridSheet(wb, "GSTR-2B", labels24, rows);
                    includedSheets.Add("GSTR-2B");
                }
                else
                {
                    skippedReasons.Add("GSTR-2B skipped (no GSTR-2B files in selection)");
                }
            }

            // 6. GSTR-3B vs GSTR-1 (12 Months: Company GSTIN | Section | Type | Apr 2025 .. Mar 2026 | Total)
            if (settings.Include3BVs1)
            {
                var rows = data.Gstr3BVs1Rows.Where(rowFilter).ToList();
                if (data.Has3BData && data.Has1Data && rows.Count > 0)
                {
                    BuildMonthGridSheet(wb, "GSTR-3B vs GSTR-1", labels12, rows);
                    includedSheets.Add("GSTR-3B vs GSTR-1");
                }
                else
                {
                    skippedReasons.Add("GSTR-3B vs GSTR-1 skipped (requires both GSTR-3B and GSTR-1 data)");
                }
            }

            // 7. GSTR-3B vs GSTR-2A (12 Months: Company GSTIN | Section | Type | Apr 2025 .. Mar 2026 | Total)
            if (settings.Include3BVs2A)
            {
                var rows = data.Gstr3BVs2ARows.Where(rowFilter).ToList();
                if (data.Has3BData && data.Has2AData && rows.Count > 0)
                {
                    BuildMonthGridSheet(wb, "GSTR-3B vs GSTR-2A", labels12, rows);
                    includedSheets.Add("GSTR-3B vs GSTR-2A");
                }
                else
                {
                    skippedReasons.Add("GSTR-3B vs GSTR-2A skipped (requires both GSTR-3B and GSTR-2A data)");
                }
            }

            // 8. GSTR-3B vs GSTR-2B (12 Months: Company GSTIN | Section | Type | Apr 2025 .. Mar 2026 | Total)
            if (settings.Include3BVs2B)
            {
                var rows = data.Gstr3BVs2BRows.Where(rowFilter).ToList();
                if (data.Has3BData && data.Has2BData && rows.Count > 0)
                {
                    BuildMonthGridSheet(wb, "GSTR-3B vs GSTR-2B", labels12, rows);
                    includedSheets.Add("GSTR-3B vs GSTR-2B");
                }
                else
                {
                    skippedReasons.Add("GSTR-3B vs GSTR-2B skipped (requires both GSTR-3B and GSTR-2B data)");
                }
            }

            if (wb.Worksheets.Count == 0)
            {
                BuildOverviewSheet(wb, data, singleGstinFilter);
                includedSheets.Add("Overview");
            }

            string safeCompany = SanitizeFileName(data.CompanyName);
            string fileName = string.IsNullOrEmpty(singleGstinFilter)
                ? $"AnnualReport-{safeCompany}-{data.FinancialYear}.xlsx"
                : $"AnnualReport-{singleGstinFilter}-{data.FinancialYear}.xlsx";

            string desiredPath = Path.Combine(outputDirectory, fileName);
            string finalPath = FileLockHelper.GetAvailableOutputPath(desiredPath);

            wb.SaveAs(finalPath);
            return finalPath;
        }

        private static void BuildOverviewSheet(
            XLWorkbook wb,
            AnnualWorkbookData data,
            string? singleGstinFilter)
        {
            var ws = wb.Worksheets.Add("Overview");
            ws.ShowGridLines = true;

            // 1. Top Card (Rows 2..10, Columns B..C) matching the screenshot 100%
            var titleRange = ws.Range("B2:C2");
            titleRange.Merge();
            ws.Cell(2, 2).Value = string.IsNullOrWhiteSpace(data.FirmHeaderTitle)
                ? "MAP & Associates"
                : data.FirmHeaderTitle;
            titleRange.Style.Font.FontName = "Calibri";
            titleRange.Style.Font.FontSize = 20;
            titleRange.Style.Font.Bold = true;
            titleRange.Style.Font.FontColor = XLColor.White;
            titleRange.Style.Fill.BackgroundColor = HeaderBlueFill;
            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(2).Height = 32;

            // Row 3: Blank gap row
            ws.Row(3).Height = 8;

            // Row 4: Thin blue separator bar across B4:C4
            var topBarRange = ws.Range("B4:C4");
            topBarRange.Merge();
            topBarRange.Style.Fill.BackgroundColor = HeaderBlueFill;
            ws.Row(4).Height = 9;

            var summaries = data.OverviewSummaries
                .Where(s => string.IsNullOrEmpty(singleGstinFilter) ||
                            string.Equals(s.Gstin, singleGstinFilter, StringComparison.OrdinalIgnoreCase))
                .ToList();

            string gstinCountDisplay = summaries.Count == 1
                ? $"1 GSTIN ({summaries[0].GstinDisplay})"
                : $"{summaries.Count} GSTINs";

            // Rows 5..9: Key-Value Card
            (string Label, string Value)[] infoRows =
            {
                ("Company Name", data.CompanyName),
                ("Contents", "GST Returns Report"),
                ("GSTIN", gstinCountDisplay),
                ("Financial Year", data.FinancialYear),
                ("Creation Time", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture))
            };

            for (int i = 0; i < infoRows.Length; i++)
            {
                int r = 5 + i;
                ws.Row(r).Height = 20;

                var lblCell = ws.Cell(r, 2);
                lblCell.Value = infoRows[i].Label;
                lblCell.Style.Font.FontName = "Calibri";
                lblCell.Style.Font.FontSize = 11;
                lblCell.Style.Font.Bold = true;
                lblCell.Style.Font.FontColor = XLColor.White;
                lblCell.Style.Fill.BackgroundColor = HeaderBlueFill;
                lblCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                var valCell = ws.Cell(r, 3);
                valCell.Value = infoRows[i].Value;
                valCell.Style.Font.FontName = "Calibri";
                valCell.Style.Font.FontSize = 11;
                valCell.Style.Font.FontColor = XLColor.Black;
                valCell.Style.Fill.BackgroundColor = XLColor.White;
                valCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                valCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                valCell.Style.Border.OutsideBorderColor = LightBorderColor;
            }

            // Row 10: Bottom Blue Bar ("Created by Octa GST")
            var bottomBarRange = ws.Range("B10:C10");
            bottomBarRange.Merge();
            ws.Cell(10, 2).Value = "Created by Octa GST";
            bottomBarRange.Style.Font.FontName = "Calibri";
            bottomBarRange.Style.Font.FontSize = 9;
            bottomBarRange.Style.Font.Bold = true;
            bottomBarRange.Style.Font.FontColor = XLColor.White;
            bottomBarRange.Style.Fill.BackgroundColor = HeaderBlueFill;
            bottomBarRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            bottomBarRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(10).Height = 18;

            // 2. Enhanced Executive Summary & Return Coverage Table below Row 13
            if (summaries.Count > 0)
            {
                ws.Cell(13, 2).Value = "Executive Summary — GSTIN-wise Return Coverage & Annual Totals";
                ws.Cell(13, 2).Style.Font.FontName = "Calibri";
                ws.Cell(13, 2).Style.Font.FontSize = 11;
                ws.Cell(13, 2).Style.Font.Bold = true;
                ws.Cell(13, 2).Style.Font.FontColor = HeaderBlueFill;

                string[] tableHeaders =
                {
                    "Company GSTIN",
                    "GSTR-3B Filed",
                    "GSTR-1 Filed",
                    "GSTR-2A Filed",
                    "GSTR-2B Filed",
                    "3B Outward Taxable",
                    "1 Outward Taxable",
                    "Outward Taxable Diff",
                    "3B Outward Tax",
                    "1 Outward Tax",
                    "Outward Tax Diff",
                    "3B Net ITC",
                    "2A Eligible ITC",
                    "2B Eligible ITC",
                    "Tax Paid via ITC",
                    "Tax Paid in Cash"
                };

                int hdrRow = 14;
                for (int c = 0; c < tableHeaders.Length; c++)
                {
                    var cell = ws.Cell(hdrRow, c + 2);
                    cell.Value = tableHeaders[c];
                    cell.Style.Font.FontName = "Calibri";
                    cell.Style.Font.FontSize = 10;
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Fill.BackgroundColor = HeaderBlueFill;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Alignment.WrapText = true;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.OutsideBorderColor = XLColor.White;
                }
                ws.Row(hdrRow).Height = 26;

                int rIdx = hdrRow + 1;
                foreach (var item in summaries)
                {
                    ws.Cell(rIdx, 2).Value = item.GstinDisplay;
                    ws.Cell(rIdx, 3).Value = $"{item.R3BMonthsCount} Months";
                    ws.Cell(rIdx, 4).Value = $"{item.R1MonthsCount} Months";
                    ws.Cell(rIdx, 5).Value = $"{item.R2AMonthsCount} Months";
                    ws.Cell(rIdx, 6).Value = $"{item.R2BMonthsCount} Months";

                    ws.Cell(rIdx, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(rIdx, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(rIdx, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(rIdx, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    SetNumCell(ws.Cell(rIdx, 7), item.OutwardTaxable3B);
                    SetNumCell(ws.Cell(rIdx, 8), item.OutwardTaxable1);
                    SetNumCell(ws.Cell(rIdx, 9), item.OutwardTaxable3B - item.OutwardTaxable1);
                    SetNumCell(ws.Cell(rIdx, 10), item.OutwardTax3B);
                    SetNumCell(ws.Cell(rIdx, 11), item.OutwardTax1);
                    SetNumCell(ws.Cell(rIdx, 12), item.OutwardTax3B - item.OutwardTax1);
                    SetNumCell(ws.Cell(rIdx, 13), item.NetItc3B);
                    SetNumCell(ws.Cell(rIdx, 14), item.Itc2A);
                    SetNumCell(ws.Cell(rIdx, 15), item.Itc2B);
                    SetNumCell(ws.Cell(rIdx, 16), item.ItcPaid3B);
                    SetNumCell(ws.Cell(rIdx, 17), item.CashPaid3B);

                    for (int c = 2; c <= 17; c++)
                    {
                        ws.Cell(rIdx, c).Style.Font.FontName = "Calibri";
                        ws.Cell(rIdx, c).Style.Font.FontSize = 9;
                        ws.Cell(rIdx, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Cell(rIdx, c).Style.Border.OutsideBorderColor = LightBorderColor;
                    }
                    rIdx++;
                }

                // Total Row
                ws.Cell(rIdx, 2).Value = "Total (All GSTINs)";
                SetNumCell(ws.Cell(rIdx, 7), summaries.Sum(x => x.OutwardTaxable3B));
                SetNumCell(ws.Cell(rIdx, 8), summaries.Sum(x => x.OutwardTaxable1));
                SetNumCell(ws.Cell(rIdx, 9), summaries.Sum(x => x.OutwardTaxable3B - x.OutwardTaxable1));
                SetNumCell(ws.Cell(rIdx, 10), summaries.Sum(x => x.OutwardTax3B));
                SetNumCell(ws.Cell(rIdx, 11), summaries.Sum(x => x.OutwardTax1));
                SetNumCell(ws.Cell(rIdx, 12), summaries.Sum(x => x.OutwardTax3B - x.OutwardTax1));
                SetNumCell(ws.Cell(rIdx, 13), summaries.Sum(x => x.NetItc3B));
                SetNumCell(ws.Cell(rIdx, 14), summaries.Sum(x => x.Itc2A));
                SetNumCell(ws.Cell(rIdx, 15), summaries.Sum(x => x.Itc2B));
                SetNumCell(ws.Cell(rIdx, 16), summaries.Sum(x => x.ItcPaid3B));
                SetNumCell(ws.Cell(rIdx, 17), summaries.Sum(x => x.CashPaid3B));

                for (int c = 2; c <= 17; c++)
                {
                    var cell = ws.Cell(rIdx, c);
                    cell.Style.Font.FontName = "Calibri";
                    cell.Style.Font.FontSize = 9;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = SectionSummaryBlueFill;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.OutsideBorderColor = LightBorderColor;
                }
            }

            // Column widths matching screenshot proportions
            ws.Column(1).Width = 4;
            ws.Column(2).Width = 22;
            ws.Column(3).Width = 52;
            for (int c = 4; c <= 6; c++) ws.Column(c).Width = 15;
            for (int c = 7; c <= 17; c++) ws.Column(c).Width = 17;
        }

        private static void BuildMonthGridSheet(
            XLWorkbook wb,
            string sheetName,
            string[] monthLabels,
            List<AnnualSheetRow> rows)
        {
            var ws = wb.Worksheets.Add(sheetName);
            int monthCount = monthLabels.Length; // 12 or 24
            int totalCol = 4 + monthCount;       // 16 (Col P) or 28 (Col AB)

            // Exact Header Names requested by User: Company GSTIN | Section | Type | Month 1 .. Month N | Total
            ws.Cell(1, 1).Value = "Company GSTIN";
            ws.Cell(1, 2).Value = "Section";
            ws.Cell(1, 3).Value = "Type";

            for (int m = 0; m < monthCount; m++)
            {
                ws.Cell(1, 4 + m).Value = monthLabels[m];
            }
            ws.Cell(1, totalCol).Value = "Total";

            StyleHeaderRow(ws, totalCol);

            string lastMonthColLetter = ws.Column(3 + monthCount).ColumnLetter();
            int rIdx = 2;
            foreach (var row in rows)
            {
                ws.Cell(rIdx, 1).Value = row.GstinDisplay;
                ws.Cell(rIdx, 2).Value = row.Section;
                ws.Cell(rIdx, 3).Value = row.Description;

                for (int m = 0; m < monthCount; m++)
                {
                    var cell = ws.Cell(rIdx, 4 + m);
                    bool hasData = m < row.HasMonthData.Length && row.HasMonthData[m];
                    if (hasData)
                    {
                        decimal val = m < row.MonthlyValues.Length ? row.MonthlyValues[m] : 0m;
                        SetNumCell(cell, val);
                    }
                    else
                    {
                        // Leave cell empty when that month's data was not uploaded ("jo month k data nhi ho usse empty cor dena")
                        cell.Clear(XLClearOptions.Contents);
                        cell.Style.NumberFormat.Format = AccountingNumberFormat;
                    }
                }

                // Total column sums all populated month cells
                ws.Cell(rIdx, totalCol).FormulaA1 = $"SUM(D{rIdx}:{lastMonthColLetter}{rIdx})";
                ws.Cell(rIdx, totalCol).Style.NumberFormat.Format = AccountingNumberFormat;

                StyleDataRow(ws, rIdx, totalCol, row.IsSectionSummaryRow, row.IsDifferenceRow);
                rIdx++;
            }

            ApplyStandardLayout(ws, monthCount);
        }

        private static void StyleHeaderRow(IXLWorksheet ws, int totalCols)
        {
            var headerRange = ws.Range(1, 1, 1, totalCols);
            headerRange.Style.Font.FontName = "Calibri";
            headerRange.Style.Font.FontSize = 10;
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = HeaderBlueFill;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.InsideBorderColor = XLColor.White;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.OutsideBorderColor = HeaderBlueFill;

            for (int c = 4; c <= totalCols; c++)
            {
                ws.Cell(1, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }

            ws.Row(1).Height = 28;
        }

        private static void StyleDataRow(
            IXLWorksheet ws,
            int rowIdx,
            int totalCols,
            bool isSectionSummary,
            bool isDifferenceRow)
        {
            var rowRange = ws.Range(rowIdx, 1, rowIdx, totalCols);
            rowRange.Style.Font.FontName = "Calibri";
            rowRange.Style.Font.FontSize = 9;
            rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            rowRange.Style.Border.InsideBorderColor = LightBorderColor;
            rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rowRange.Style.Border.OutsideBorderColor = LightBorderColor;

            ws.Cell(rowIdx, totalCols).Style.Font.Bold = true;

            if (isSectionSummary)
            {
                rowRange.Style.Font.Bold = true;
                rowRange.Style.Fill.BackgroundColor = SectionSummaryBlueFill;
            }
            else if (isDifferenceRow)
            {
                rowRange.Style.Font.Bold = true;
                rowRange.Style.Fill.BackgroundColor = DifferenceGrayFill;
            }
        }

        private static void ApplyStandardLayout(IXLWorksheet ws, int monthCount)
        {
            int totalCol = 4 + monthCount;
            ws.Column(1).Width = 24; // Company GSTIN
            ws.Column(2).Width = 42; // Section
            ws.Column(3).Width = 18; // Type

            for (int c = 4; c < totalCol; c++)
            {
                ws.Column(c).Width = 14;
            }
            ws.Column(totalCol).Width = 20;

            // Freeze Panes at D2 (freeze top 1 row and first 3 columns: Company GSTIN, Section, Type)
            ws.SheetView.FreezeRows(1);
            ws.SheetView.FreezeColumns(3);

            var used = ws.RangeUsed();
            if (used != null)
            {
                used.SetAutoFilter();
            }
        }

        private static void SetNumCell(IXLCell cell, decimal value)
        {
            cell.Value = value;
            cell.Style.NumberFormat.Format = AccountingNumberFormat;
        }

        private static string[] Get12MonthLabels(int startYear, int endYear)
        {
            return new[]
            {
                $"Apr {startYear}",
                $"May {startYear}",
                $"Jun {startYear}",
                $"Jul {startYear}",
                $"Aug {startYear}",
                $"Sep {startYear}",
                $"Oct {startYear}",
                $"Nov {startYear}",
                $"Dec {startYear}",
                $"Jan {endYear}",
                $"Feb {endYear}",
                $"Mar {endYear}"
            };
        }

        private static string[] Get24MonthLabels(int startYear, int endYear, int nextEndYear)
        {
            return new[]
            {
                $"Apr {startYear}",
                $"May {startYear}",
                $"Jun {startYear}",
                $"Jul {startYear}",
                $"Aug {startYear}",
                $"Sep {startYear}",
                $"Oct {startYear}",
                $"Nov {startYear}",
                $"Dec {startYear}",
                $"Jan {endYear}",
                $"Feb {endYear}",
                $"Mar {endYear}",
                $"Apr {endYear}",
                $"May {endYear}",
                $"Jun {endYear}",
                $"Jul {endYear}",
                $"Aug {endYear}",
                $"Sep {endYear}",
                $"Oct {endYear}",
                $"Nov {endYear}",
                $"Dec {endYear}",
                $"Jan {nextEndYear}",
                $"Feb {nextEndYear}",
                $"Mar {nextEndYear}"
            };
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Company";
            var invalid = Path.GetInvalidFileNameChars();
            var cleaned = new string(name.Where(ch => !invalid.Contains(ch)).ToArray()).Trim();
            return string.IsNullOrWhiteSpace(cleaned) ? "Company" : cleaned;
        }
    }
}
