using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Helpers;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class OctaGstBuilderService : IOctaGstBuilderService
    {
        private readonly ILoggingService _logger;
        private static readonly XLColor RoyalBlue = XLColor.FromHtml("#0070C0");

        public OctaGstBuilderService(ILoggingService logger)
        {
            _logger = logger;
        }

        public async Task<(bool Success, int TotalRecords, string? Error)> BuildOctaWorkbookAsync(
            GstFileType fileType,
            List<ScannedFileItem> sourceFiles,
            string outputExcelPath,
            CancellationToken cancellationToken = default)
        {
            try
            {
                using var workbook = new XLWorkbook();
                var companyGstins = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var periods = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                string companyName = "Test_SPECTAL MANAGEMENT";

                var jsonDocs = new List<(string FileName, JsonDocument Doc)>();
                var allLeafNodes = new Dictionary<string, JsonFlatValue>(StringComparer.Ordinal);

                foreach (var file in sourceFiles)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    await using var stream = File.OpenRead(file.FilePath);
                    var doc = await JsonDocument.ParseAsync(stream, new JsonDocumentOptions
                    {
                        AllowTrailingCommas = true,
                        CommentHandling = JsonCommentHandling.Skip
                    }, cancellationToken);

                    jsonDocs.Add((file.FileName, doc));

                    // Extract GSTIN and Tax Period from data or root
                    var root = doc.RootElement;
                    var data = ResolveDataElement(root);

                    string g = GetPropString(data, "gstin");
                    if (string.IsNullOrEmpty(g)) g = GetPropString(root, "gstin");
                    if (string.IsNullOrEmpty(g))
                    {
                        var mGstin = Regex.Match(file.FileName, @"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}");
                        if (mGstin.Success) g = mGstin.Value;
                    }
                    if (!string.IsNullOrEmpty(g)) companyGstins.Add(g);

                    string cn = GetPropString(data, "cname", "trade_name", "lgl_name", "trdnm");
                    if (string.IsNullOrEmpty(cn)) cn = GetPropString(root, "cname", "trade_name", "lgl_name", "trdnm");
                    if (!string.IsNullOrEmpty(cn)) companyName = cn;

                    string fp = GetPropString(data, "fp", "ret_period", "rtnprd");
                    if (string.IsNullOrEmpty(fp)) fp = GetPropString(root, "fp", "ret_period", "rtnprd");
                    if (string.IsNullOrEmpty(fp))
                    {
                        var mFp = Regex.Match(file.FileName, @"_([0-1][0-9]20[2-3][0-9])_");
                        if (mFp.Success) fp = mFp.Groups[1].Value;
                    }
                    if (!string.IsNullOrEmpty(fp)) periods.Add(fp);

                    // Index leaves for All_Data_Index safety net
                    IndexLeaves(doc.RootElement, $"[{file.FileName}] $", allLeafNodes);
                }

                if (companyGstins.Count == 0) companyGstins.Add("07ACWFS8659K2ZV");
                string periodRange = FormatPeriodRange(periods);

                // 1. Build Overview Sheet
                var overviewSheet = workbook.Worksheets.Add("Overview");
                BuildOverviewSheet(overviewSheet, fileType, companyGstins, companyName, periodRange);

                int totalRecords = 0;

                // 2. Build Type-Specific Standard Sheets
                switch (fileType)
                {
                    case GstFileType.R1:
                        totalRecords += BuildR1Sheets(workbook, jsonDocs);
                        break;

                    case GstFileType.R2A:
                        totalRecords += BuildR2ASheets(workbook, jsonDocs);
                        break;

                    case GstFileType.R2B:
                        totalRecords += BuildR2BSheets(workbook, jsonDocs);
                        break;

                    default:
                        // Fallback generic mapping
                        break;
                }

                // 3. Build All_Data_Index Sheet for 110% zero data loss guarantee
                var indexSheet = workbook.Worksheets.Add("All_Data_Index");
                BuildAuditIndexSheet(indexSheet, allLeafNodes);

                // Ensure output directory exists and save
                string? dir = Path.GetDirectoryName(outputExcelPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                outputExcelPath = FileLockHelper.GetAvailableOutputPath(outputExcelPath);
                workbook.SaveAs(outputExcelPath);
                _logger.LogInfo($"Octa format Excel '{Path.GetFileName(outputExcelPath)}' created successfully ({totalRecords} records).");

                return (true, totalRecords, null);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed building Octa GST Excel for {fileType}", ex);
                return (false, 0, ex.Message);
            }
        }

        #region Overview Sheet (Matching Reference Screenshot)

        private void BuildOverviewSheet(
            IXLWorksheet sheet,
            GstFileType fileType,
            HashSet<string> companyGstins,
            string companyName,
            string periodRange)
        {
            sheet.ShowGridLines = true;

            // Row 1: Header Banner (Royal Blue, White Bold text)
            var banner = sheet.Range("A1:C1");
            banner.Merge();
            banner.Value = "MAP & Associates";
            banner.Style.Font.Bold = true;
            banner.Style.Font.FontSize = 14;
            banner.Style.Font.FontColor = XLColor.White;
            banner.Style.Fill.BackgroundColor = RoyalBlue;
            banner.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            banner.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Row(1).Height = 28;

            sheet.Row(2).Height = 8; // Spacer

            string contentsTitle = fileType switch
            {
                GstFileType.R1 => "GSTR-1/1A Data",
                GstFileType.R2A => "GSTR-2A",
                GstFileType.R2B => "GSTR-2B Data",
                _ => "GSTR-3B Data"
            };

            int currentRow = 3;

            // Row: Company Name
            SetOverviewRow(sheet, currentRow++, "Company Name", companyName);

            // Row: Contents
            SetOverviewRow(sheet, currentRow++, "Contents", contentsTitle);

            // Rows: Company GSTIN(s) with State names
            int gstinStartRow = currentRow;
            foreach (var gstin in companyGstins)
            {
                string formatted = GstStateHelper.FormatGstinWithState(gstin);
                sheet.Cell(currentRow, 2).SetValue(formatted);
                sheet.Cell(currentRow, 2).Style.Font.FontSize = 10.5;
                sheet.Cell(currentRow, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheet.Cell(currentRow, 2).Style.Border.OutsideBorderColor = XLColor.FromHtml("#E2E8F0");
                currentRow++;
            }
            int gstinEndRow = currentRow - 1;

            if (gstinEndRow >= gstinStartRow)
            {
                var gstinLabel = sheet.Range(gstinStartRow, 1, gstinEndRow, 1);
                if (gstinStartRow != gstinEndRow) gstinLabel.Merge();
                gstinLabel.Value = "Company GSTIN";
                gstinLabel.Style.Font.Bold = true;
                gstinLabel.Style.Font.FontColor = XLColor.White;
                gstinLabel.Style.Fill.BackgroundColor = RoyalBlue;
                gstinLabel.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            // Row: Period
            SetOverviewRow(sheet, currentRow++, "Period", periodRange);

            // Row: Creation Time
            SetOverviewRow(sheet, currentRow++, "Creation Time", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));

            sheet.Row(currentRow).Height = 8; // Spacer
            currentRow++;

            // Bottom Right Footer: Created by Octa GST
            var footerRange = sheet.Range(currentRow, 1, currentRow, 3);
            footerRange.Merge();
            footerRange.Value = "Created by Octa GST";
            footerRange.Style.Font.Bold = true;
            footerRange.Style.Font.FontSize = 10;
            footerRange.Style.Font.FontColor = XLColor.White;
            footerRange.Style.Fill.BackgroundColor = RoyalBlue;
            footerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            footerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Row(currentRow).Height = 22;

            sheet.Column(1).Width = 24;
            sheet.Column(2).Width = 55;
            sheet.Column(3).Width = 15;
        }

        private static void SetOverviewRow(IXLWorksheet sheet, int row, string label, string value)
        {
            var cellA = sheet.Cell(row, 1);
            cellA.SetValue(label);
            cellA.Style.Font.Bold = true;
            cellA.Style.Font.FontColor = XLColor.White;
            cellA.Style.Fill.BackgroundColor = RoyalBlue;
            cellA.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            var cellB = sheet.Cell(row, 2);
            cellB.SetValue(value);
            cellB.Style.Font.FontSize = 10.5;
            cellB.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cellB.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cellB.Style.Border.OutsideBorderColor = XLColor.FromHtml("#E2E8F0");

            sheet.Row(row).Height = 20;
        }

        #endregion

        #region R1 (GSTR-1) Sheets

        private int BuildR1Sheets(XLWorkbook wb, List<(string FileName, JsonDocument Doc)> docs)
        {
            int totalRecords = 0;

            // Sheet: Sales
            string[] salesCols = {
                "Company GSTIN", "Tax Period", "Doc Type", "Sale Type", "Doc No", "Doc Date",
                "Customer GSTIN", "Customer Name", "Place of Supply", "Shipping Bill Date",
                "Shipping Bill No", "Port Code", "Reference Doc No", "Reference Doc Date",
                "Reverse Charge", "Doc Value", "Item Taxable Value", "GST Rate", "IGST",
                "CGST", "SGST", "Cess", "Is Amendment", "Original Doc No", "Original Doc Date",
                "Uploaded By", "Source", "IRN Date", "IRN", "GSTR-1A"
            };
            var salesSheet = wb.Worksheets.Add("Sales");
            SetupTableHeaders(salesSheet, salesCols);
            int salesRowIdx = 2;

            // Sheet: Sales Summary
            string[] summaryCols = {
                "Company GSTIN", "Tax Period", "Summary Type", "Description", "Place of Supply",
                "Taxable Value", "GST Rate", "IGST", "CGST", "SGST", "Cess", "Ecommerce GSTIN",
                "Original Period", "Original Ecommerce GSTIN", "GSTR-1A"
            };
            var summarySheet = wb.Worksheets.Add("Sales Summary");
            SetupTableHeaders(summarySheet, summaryCols);
            int summaryRowIdx = 2;

            // Sheet: SalesHSN
            string[] hsnCols = {
                "Company GSTIN", "Tax Period", "Summary Type", "HSN Code", "HSN Description",
                "UQC", "Qty", "Total Value", "Taxable Value", "GST Rate", "IGST", "CGST",
                "SGST", "Cess", "GSTR-1A"
            };
            var hsnSheet = wb.Worksheets.Add("SalesHSN");
            SetupTableHeaders(hsnSheet, hsnCols);
            int hsnRowIdx = 2;

            // Sheet: Disclosed
            string[] discCols = {
                "Company GSTIN", "Tax Period", "Doc Type", "From", "To", "Total", "Cancelled", "GSTR-1A"
            };
            var discSheet = wb.Worksheets.Add("Disclosed");
            SetupTableHeaders(discSheet, discCols);
            int discRowIdx = 2;

            foreach (var (fileName, doc) in docs)
            {
                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object) continue;

                var data = ResolveDataElement(root);

                string companyGstin = GetPropString(data, "gstin");
                if (string.IsNullOrEmpty(companyGstin)) companyGstin = GetPropString(root, "gstin");
                if (string.IsNullOrEmpty(companyGstin))
                {
                    var mGstin = Regex.Match(fileName, @"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}");
                    if (mGstin.Success) companyGstin = mGstin.Value;
                }
                if (string.IsNullOrEmpty(companyGstin)) companyGstin = "07ACWFS8659K2ZV";

                string rawFp = GetPropString(data, "fp", "ret_period", "rtnprd");
                if (string.IsNullOrEmpty(rawFp)) rawFp = GetPropString(root, "fp", "ret_period", "rtnprd");
                if (string.IsNullOrEmpty(rawFp))
                {
                    var mFp = Regex.Match(fileName, @"_([0-1][0-9]20[2-3][0-9])_");
                    if (mFp.Success) rawFp = mFp.Groups[1].Value;
                }
                string taxPeriod = FormatSinglePeriod(rawFp);

                JsonElement r1Data = data.TryGetProperty("b2b", out _) || data.TryGetProperty("hsn", out _) ? data : root;

                // 1. Process B2B / B2BA
                ProcessR1InvoiceArray(r1Data, "b2b", "B2B", "Invoice", "No", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);
                ProcessR1InvoiceArray(r1Data, "b2ba", "B2B", "Invoice", "Yes", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);

                // 2. Process B2CL / B2CLA
                ProcessR1InvoiceArray(r1Data, "b2cl", "B2CL", "Invoice", "No", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);
                ProcessR1InvoiceArray(r1Data, "b2cla", "B2CL", "Invoice", "Yes", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);

                // 3. Process CDNR / CDNRA
                ProcessR1CdnrArray(r1Data, "cdnr", "CDNR", "No", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);
                ProcessR1CdnrArray(r1Data, "cdnra", "CDNR", "Yes", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);

                // 4. Process CDNUR / CDNURA
                ProcessR1CdnrArray(r1Data, "cdnur", "CDNUR", "No", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);
                ProcessR1CdnrArray(r1Data, "cdnura", "CDNUR", "Yes", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);

                // 5. Process EXP / EXPA
                ProcessR1ExpArray(r1Data, "exp", "EXP", "No", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);
                ProcessR1ExpArray(r1Data, "expa", "EXP", "Yes", companyGstin, taxPeriod, salesSheet, ref salesRowIdx);

                // 6. Process Sales Summary (B2CS, NIL, AT, TXPD)
                ProcessR1B2cs(r1Data, companyGstin, taxPeriod, summarySheet, ref summaryRowIdx);
                ProcessR1Nil(r1Data, companyGstin, taxPeriod, summarySheet, ref summaryRowIdx);

                // 7. Process HSN
                ProcessR1Hsn(r1Data, companyGstin, taxPeriod, hsnSheet, ref hsnRowIdx);

                // 8. Process Disclosed (Doc Issue)
                ProcessR1DocIssue(r1Data, companyGstin, taxPeriod, discSheet, ref discRowIdx);
            }

            FinalizeTableLayout(salesSheet, salesCols.Length, salesRowIdx);
            FinalizeTableLayout(summarySheet, summaryCols.Length, summaryRowIdx);
            FinalizeTableLayout(hsnSheet, hsnCols.Length, hsnRowIdx);
            FinalizeTableLayout(discSheet, discCols.Length, discRowIdx);

            totalRecords = (salesRowIdx - 2) + (summaryRowIdx - 2) + (hsnRowIdx - 2) + (discRowIdx - 2);
            return totalRecords;
        }

        private static void ProcessR1InvoiceArray(
            JsonElement root, string propName, string saleType, string defaultDocType, string isAmendment,
            string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var party in arr.EnumerateArray())
            {
                string customerGstin = GetPropString(party, "ctin");
                string customerName = GetPropString(party, "cname", "trade_name");

                if (party.TryGetProperty("inv", out var invArr) && invArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var inv in invArr.EnumerateArray())
                    {
                        string docNo = GetPropString(inv, "inum");
                        string docDate = GetPropString(inv, "idt");
                        string pos = GetPropString(inv, "pos");
                        string posFormatted = FormatPos(pos);
                        string rchrg = GetPropString(inv, "rchrg", "rev");
                        if (string.IsNullOrEmpty(rchrg)) rchrg = "N";
                        decimal docVal = GetPropDecimal(inv, "val");
                        string origDocNo = GetPropString(inv, "oinum");
                        string origDocDate = GetPropString(inv, "oidt");
                        string irn = GetPropString(inv, "irn");
                        string irnDate = GetPropString(inv, "irngendate", "irn_dt");
                        string updby = GetPropString(inv, "updby");
                        string source = !string.IsNullOrEmpty(irn) ? "E-Invoice" : "GSTR-1";

                        if (inv.TryGetProperty("itms", out var itmsArr) && itmsArr.ValueKind == JsonValueKind.Array && itmsArr.GetArrayLength() > 0)
                        {
                            foreach (var itm in itmsArr.EnumerateArray())
                            {
                                JsonElement det = itm.TryGetProperty("itm_det", out var d) ? d : itm;

                                decimal txval = GetPropDecimal(det, "txval");
                                decimal rt = GetPropDecimal(det, "rt");
                                decimal iamt = GetPropDecimal(det, "iamt");
                                decimal camt = GetPropDecimal(det, "camt");
                                decimal samt = GetPropDecimal(det, "samt");
                                decimal csamt = GetPropDecimal(det, "csamt");
                                if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                                sheet.Cell(rowIdx, 3).SetValue(defaultDocType);
                                sheet.Cell(rowIdx, 4).SetValue(saleType);
                                SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                                sheet.Cell(rowIdx, 6).SetValue(docDate);
                                SetTextCell(sheet.Cell(rowIdx, 7), customerGstin);
                                sheet.Cell(rowIdx, 8).SetValue(customerName);
                                sheet.Cell(rowIdx, 9).SetValue(posFormatted);
                                sheet.Cell(rowIdx, 15).SetValue(rchrg);
                                SetNumberCell(sheet.Cell(rowIdx, 16), docVal);
                                SetNumberCell(sheet.Cell(rowIdx, 17), txval);
                                SetNumberCell(sheet.Cell(rowIdx, 18), rt);
                                SetNumberCell(sheet.Cell(rowIdx, 19), iamt);
                                SetNumberCell(sheet.Cell(rowIdx, 20), camt);
                                SetNumberCell(sheet.Cell(rowIdx, 21), samt);
                                SetNumberCell(sheet.Cell(rowIdx, 22), csamt);
                                sheet.Cell(rowIdx, 23).SetValue(isAmendment);
                                SetTextCell(sheet.Cell(rowIdx, 24), origDocNo);
                                sheet.Cell(rowIdx, 25).SetValue(origDocDate);
                                sheet.Cell(rowIdx, 26).SetValue(updby);
                                sheet.Cell(rowIdx, 27).SetValue(source);
                                sheet.Cell(rowIdx, 28).SetValue(irnDate);
                                SetTextCell(sheet.Cell(rowIdx, 29), irn);
                                sheet.Cell(rowIdx, 30).SetValue("N");

                                rowIdx++;
                            }
                        }
                        else
                        {
                            decimal txval = GetPropDecimal(inv, "txval");
                            decimal rt = GetPropDecimal(inv, "rt");
                            decimal iamt = GetPropDecimal(inv, "iamt", "igst");
                            decimal camt = GetPropDecimal(inv, "camt", "cgst");
                            decimal samt = GetPropDecimal(inv, "samt", "sgst");
                            decimal csamt = GetPropDecimal(inv, "csamt", "cess");
                            if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                            sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                            sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                            sheet.Cell(rowIdx, 3).SetValue(defaultDocType);
                            sheet.Cell(rowIdx, 4).SetValue(saleType);
                            SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                            sheet.Cell(rowIdx, 6).SetValue(docDate);
                            SetTextCell(sheet.Cell(rowIdx, 7), customerGstin);
                            sheet.Cell(rowIdx, 8).SetValue(customerName);
                            sheet.Cell(rowIdx, 9).SetValue(posFormatted);
                            sheet.Cell(rowIdx, 15).SetValue(rchrg);
                            SetNumberCell(sheet.Cell(rowIdx, 16), docVal);
                            SetNumberCell(sheet.Cell(rowIdx, 17), txval);
                            SetNumberCell(sheet.Cell(rowIdx, 18), rt);
                            SetNumberCell(sheet.Cell(rowIdx, 19), iamt);
                            SetNumberCell(sheet.Cell(rowIdx, 20), camt);
                            SetNumberCell(sheet.Cell(rowIdx, 21), samt);
                            SetNumberCell(sheet.Cell(rowIdx, 22), csamt);
                            sheet.Cell(rowIdx, 23).SetValue(isAmendment);
                            SetTextCell(sheet.Cell(rowIdx, 24), origDocNo);
                            sheet.Cell(rowIdx, 25).SetValue(origDocDate);
                            sheet.Cell(rowIdx, 26).SetValue(updby);
                            sheet.Cell(rowIdx, 27).SetValue(source);
                            sheet.Cell(rowIdx, 28).SetValue(irnDate);
                            SetTextCell(sheet.Cell(rowIdx, 29), irn);
                            sheet.Cell(rowIdx, 30).SetValue("N");

                            rowIdx++;
                        }
                    }
                }
            }
        }

        private static void ProcessR1CdnrArray(
            JsonElement root, string propName, string saleType, string isAmendment,
            string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var party in arr.EnumerateArray())
            {
                string customerGstin = GetPropString(party, "ctin");
                string customerName = GetPropString(party, "cname", "trade_name");

                var ntArr = party.TryGetProperty("nt", out var nts) ? nts : (party.TryGetProperty("inv", out var invs) ? invs : default);
                if (ntArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var nt in ntArr.EnumerateArray())
                    {
                        string ntty = GetPropString(nt, "ntty", "typ");
                        string docType = (ntty == "C" || ntty.Equals("Credit", StringComparison.OrdinalIgnoreCase)) ? "Credit Note"
                                       : ((ntty == "D" || ntty.Equals("Debit", StringComparison.OrdinalIgnoreCase)) ? "Debit Note" : "Credit Note");
                        string docNo = GetPropString(nt, "nt_num", "ntnum", "inum");
                        string docDate = GetPropString(nt, "nt_dt", "dt", "idt");
                        string refDocNo = GetPropString(nt, "inum", "p_gst");
                        string refDocDate = GetPropString(nt, "idt");
                        decimal docVal = GetPropDecimal(nt, "val");
                        string pos = GetPropString(nt, "pos");
                        string posFormatted = FormatPos(pos);
                        string rchrg = GetPropString(nt, "rchrg", "rev");
                        if (string.IsNullOrEmpty(rchrg)) rchrg = "N";
                        string origDocNo = GetPropString(nt, "ont_num", "oinum");
                        string origDocDate = GetPropString(nt, "ont_dt", "oidt");
                        string irn = GetPropString(nt, "irn");
                        string irnDate = GetPropString(nt, "irngendate", "irn_dt");
                        string updby = GetPropString(nt, "updby");
                        string source = !string.IsNullOrEmpty(irn) ? "E-Invoice" : "GSTR-1";

                        if (nt.TryGetProperty("itms", out var itmsArr) && itmsArr.ValueKind == JsonValueKind.Array && itmsArr.GetArrayLength() > 0)
                        {
                            foreach (var itm in itmsArr.EnumerateArray())
                            {
                                JsonElement det = itm.TryGetProperty("itm_det", out var d) ? d : itm;

                                decimal txval = GetPropDecimal(det, "txval");
                                decimal rt = GetPropDecimal(det, "rt");
                                decimal iamt = GetPropDecimal(det, "iamt");
                                decimal camt = GetPropDecimal(det, "camt");
                                decimal samt = GetPropDecimal(det, "samt");
                                decimal csamt = GetPropDecimal(det, "csamt");
                                if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                                sheet.Cell(rowIdx, 3).SetValue(docType);
                                sheet.Cell(rowIdx, 4).SetValue(saleType);
                                SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                                sheet.Cell(rowIdx, 6).SetValue(docDate);
                                SetTextCell(sheet.Cell(rowIdx, 7), customerGstin);
                                sheet.Cell(rowIdx, 8).SetValue(customerName);
                                sheet.Cell(rowIdx, 9).SetValue(posFormatted);
                                SetTextCell(sheet.Cell(rowIdx, 13), refDocNo);
                                sheet.Cell(rowIdx, 14).SetValue(refDocDate);
                                sheet.Cell(rowIdx, 15).SetValue(rchrg);
                                SetNumberCell(sheet.Cell(rowIdx, 16), docVal);
                                SetNumberCell(sheet.Cell(rowIdx, 17), txval);
                                SetNumberCell(sheet.Cell(rowIdx, 18), rt);
                                SetNumberCell(sheet.Cell(rowIdx, 19), iamt);
                                SetNumberCell(sheet.Cell(rowIdx, 20), camt);
                                SetNumberCell(sheet.Cell(rowIdx, 21), samt);
                                SetNumberCell(sheet.Cell(rowIdx, 22), csamt);
                                sheet.Cell(rowIdx, 23).SetValue(isAmendment);
                                SetTextCell(sheet.Cell(rowIdx, 24), origDocNo);
                                sheet.Cell(rowIdx, 25).SetValue(origDocDate);
                                sheet.Cell(rowIdx, 26).SetValue(updby);
                                sheet.Cell(rowIdx, 27).SetValue(source);
                                sheet.Cell(rowIdx, 28).SetValue(irnDate);
                                SetTextCell(sheet.Cell(rowIdx, 29), irn);
                                sheet.Cell(rowIdx, 30).SetValue("N");

                                rowIdx++;
                            }
                        }
                        else
                        {
                            decimal txval = GetPropDecimal(nt, "txval");
                            decimal rt = GetPropDecimal(nt, "rt");
                            decimal iamt = GetPropDecimal(nt, "iamt", "igst");
                            decimal camt = GetPropDecimal(nt, "camt", "cgst");
                            decimal samt = GetPropDecimal(nt, "samt", "sgst");
                            decimal csamt = GetPropDecimal(nt, "csamt", "cess");
                            if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                            sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                            sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                            sheet.Cell(rowIdx, 3).SetValue(docType);
                            sheet.Cell(rowIdx, 4).SetValue(saleType);
                            SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                            sheet.Cell(rowIdx, 6).SetValue(docDate);
                            SetTextCell(sheet.Cell(rowIdx, 7), customerGstin);
                            sheet.Cell(rowIdx, 8).SetValue(customerName);
                            sheet.Cell(rowIdx, 9).SetValue(posFormatted);
                            SetTextCell(sheet.Cell(rowIdx, 13), refDocNo);
                            sheet.Cell(rowIdx, 14).SetValue(refDocDate);
                            sheet.Cell(rowIdx, 15).SetValue(rchrg);
                            SetNumberCell(sheet.Cell(rowIdx, 16), docVal);
                            SetNumberCell(sheet.Cell(rowIdx, 17), txval);
                            SetNumberCell(sheet.Cell(rowIdx, 18), rt);
                            SetNumberCell(sheet.Cell(rowIdx, 19), iamt);
                            SetNumberCell(sheet.Cell(rowIdx, 20), camt);
                            SetNumberCell(sheet.Cell(rowIdx, 21), samt);
                            SetNumberCell(sheet.Cell(rowIdx, 22), csamt);
                            sheet.Cell(rowIdx, 23).SetValue(isAmendment);
                            SetTextCell(sheet.Cell(rowIdx, 24), origDocNo);
                            sheet.Cell(rowIdx, 25).SetValue(origDocDate);
                            sheet.Cell(rowIdx, 26).SetValue(updby);
                            sheet.Cell(rowIdx, 27).SetValue(source);
                            sheet.Cell(rowIdx, 28).SetValue(irnDate);
                            SetTextCell(sheet.Cell(rowIdx, 29), irn);
                            sheet.Cell(rowIdx, 30).SetValue("N");

                            rowIdx++;
                        }
                    }
                }
            }
        }

        private static void ProcessR1ExpArray(
            JsonElement root, string propName, string saleType, string isAmendment,
            string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var exp in arr.EnumerateArray())
            {
                if (exp.TryGetProperty("inv", out var invArr) && invArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var inv in invArr.EnumerateArray())
                    {
                        string docNo = GetPropString(inv, "inum");
                        string docDate = GetPropString(inv, "idt");
                        decimal docVal = GetPropDecimal(inv, "val");
                        string sbnum = GetPropString(inv, "sbnum");
                        string sbdt = GetPropString(inv, "sbdt");
                        string sbpcode = GetPropString(inv, "sbpcode", "port_code");
                        string irn = GetPropString(inv, "irn");
                        string irnDate = GetPropString(inv, "irngendate", "irn_dt");
                        string updby = GetPropString(inv, "updby");
                        string source = !string.IsNullOrEmpty(irn) ? "E-Invoice" : "GSTR-1";

                        if (inv.TryGetProperty("itms", out var itmsArr) && itmsArr.ValueKind == JsonValueKind.Array && itmsArr.GetArrayLength() > 0)
                        {
                            foreach (var itm in itmsArr.EnumerateArray())
                            {
                                JsonElement det = itm.TryGetProperty("itm_det", out var d) ? d : itm;
                                decimal txval = GetPropDecimal(det, "txval");
                                decimal rt = GetPropDecimal(det, "rt");
                                decimal iamt = GetPropDecimal(det, "iamt");
                                decimal csamt = GetPropDecimal(det, "csamt");
                                if (rt == 0 && txval > 0 && iamt > 0) rt = Math.Round(iamt * 100m / txval, 2);

                                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                                sheet.Cell(rowIdx, 3).SetValue("Export Invoice");
                                sheet.Cell(rowIdx, 4).SetValue(saleType);
                                SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                                sheet.Cell(rowIdx, 6).SetValue(docDate);
                                sheet.Cell(rowIdx, 10).SetValue(sbdt);
                                SetTextCell(sheet.Cell(rowIdx, 11), sbnum);
                                SetTextCell(sheet.Cell(rowIdx, 12), sbpcode);
                                SetNumberCell(sheet.Cell(rowIdx, 16), docVal);
                                SetNumberCell(sheet.Cell(rowIdx, 17), txval);
                                SetNumberCell(sheet.Cell(rowIdx, 18), rt);
                                SetNumberCell(sheet.Cell(rowIdx, 19), iamt);
                                SetNumberCell(sheet.Cell(rowIdx, 22), csamt);
                                sheet.Cell(rowIdx, 23).SetValue(isAmendment);
                                sheet.Cell(rowIdx, 26).SetValue(updby);
                                sheet.Cell(rowIdx, 27).SetValue(source);
                                sheet.Cell(rowIdx, 28).SetValue(irnDate);
                                SetTextCell(sheet.Cell(rowIdx, 29), irn);
                                sheet.Cell(rowIdx, 30).SetValue("N");

                                rowIdx++;
                            }
                        }
                        else
                        {
                            decimal txval = GetPropDecimal(inv, "txval");
                            decimal rt = GetPropDecimal(inv, "rt");
                            decimal iamt = GetPropDecimal(inv, "iamt", "igst");
                            decimal csamt = GetPropDecimal(inv, "csamt", "cess");
                            if (rt == 0 && txval > 0 && iamt > 0) rt = Math.Round(iamt * 100m / txval, 2);

                            sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                            sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                            sheet.Cell(rowIdx, 3).SetValue("Export Invoice");
                            sheet.Cell(rowIdx, 4).SetValue(saleType);
                            SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                            sheet.Cell(rowIdx, 6).SetValue(docDate);
                            sheet.Cell(rowIdx, 10).SetValue(sbdt);
                            SetTextCell(sheet.Cell(rowIdx, 11), sbnum);
                            SetTextCell(sheet.Cell(rowIdx, 12), sbpcode);
                            SetNumberCell(sheet.Cell(rowIdx, 16), docVal);
                            SetNumberCell(sheet.Cell(rowIdx, 17), txval);
                            SetNumberCell(sheet.Cell(rowIdx, 18), rt);
                            SetNumberCell(sheet.Cell(rowIdx, 19), iamt);
                            SetNumberCell(sheet.Cell(rowIdx, 22), csamt);
                            sheet.Cell(rowIdx, 23).SetValue(isAmendment);
                            sheet.Cell(rowIdx, 26).SetValue(updby);
                            sheet.Cell(rowIdx, 27).SetValue(source);
                            sheet.Cell(rowIdx, 28).SetValue(irnDate);
                            SetTextCell(sheet.Cell(rowIdx, 29), irn);
                            sheet.Cell(rowIdx, 30).SetValue("N");

                            rowIdx++;
                        }
                    }
                }
            }
        }

        private static void ProcessR1B2cs(JsonElement root, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty("b2cs", out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var item in arr.EnumerateArray())
            {
                string pos = GetPropString(item, "pos");
                decimal txval = GetPropDecimal(item, "txval");
                decimal rt = GetPropDecimal(item, "rt");
                decimal iamt = GetPropDecimal(item, "iamt");
                decimal camt = GetPropDecimal(item, "camt");
                decimal samt = GetPropDecimal(item, "samt");
                decimal csamt = GetPropDecimal(item, "csamt");
                string etin = GetPropString(item, "etin");

                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                sheet.Cell(rowIdx, 3).SetValue("B2CS");
                sheet.Cell(rowIdx, 4).SetValue("B2C (Others)");
                sheet.Cell(rowIdx, 5).SetValue(pos);
                SetNumberCell(sheet.Cell(rowIdx, 6), txval);
                SetNumberCell(sheet.Cell(rowIdx, 7), rt);
                SetNumberCell(sheet.Cell(rowIdx, 8), iamt);
                SetNumberCell(sheet.Cell(rowIdx, 9), camt);
                SetNumberCell(sheet.Cell(rowIdx, 10), samt);
                SetNumberCell(sheet.Cell(rowIdx, 11), csamt);
                SetTextCell(sheet.Cell(rowIdx, 12), etin);
                sheet.Cell(rowIdx, 15).SetValue("N");

                rowIdx++;
            }
        }

        private static void ProcessR1Nil(JsonElement root, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty("nil", out var nilObj) || nilObj.ValueKind != JsonValueKind.Object) return;

            if (nilObj.TryGetProperty("inv", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in arr.EnumerateArray())
                {
                    decimal nilAmt = GetPropDecimal(item, "nil_amt");
                    decimal exptAmt = GetPropDecimal(item, "expt_amt");
                    decimal ngsupAmt = GetPropDecimal(item, "ngsup_amt");
                    string splyTy = GetPropString(item, "sply_ty");

                    sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                    sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                    sheet.Cell(rowIdx, 3).SetValue("NIL");
                    sheet.Cell(rowIdx, 4).SetValue($"Nil/Exempt ({splyTy})");
                    SetNumberCell(sheet.Cell(rowIdx, 6), nilAmt + exptAmt + ngsupAmt);
                    SetNumberCell(sheet.Cell(rowIdx, 7), 0);
                    sheet.Cell(rowIdx, 15).SetValue("N");

                    rowIdx++;
                }
            }
        }

        private static void ProcessR1Hsn(JsonElement root, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty("hsn", out var hsnObj)) return;

            JsonElement arr = default;
            if (hsnObj.ValueKind == JsonValueKind.Array) arr = hsnObj;
            else if (hsnObj.ValueKind == JsonValueKind.Object)
            {
                if (hsnObj.TryGetProperty("hsn_b2b", out var hb) && hb.ValueKind == JsonValueKind.Array) arr = hb;
                else if (hsnObj.TryGetProperty("data", out var d) && d.ValueKind == JsonValueKind.Array) arr = d;
            }

            if (arr.ValueKind != JsonValueKind.Array) return;

            foreach (var item in arr.EnumerateArray())
            {
                string hsnSc = GetPropString(item, "hsn_sc");
                string desc = GetPropString(item, "desc");
                string uqc = GetPropString(item, "uqc");
                decimal qty = GetPropDecimal(item, "qty");
                decimal val = GetPropDecimal(item, "val");
                decimal txval = GetPropDecimal(item, "txval");
                decimal rt = GetPropDecimal(item, "rt");
                decimal iamt = GetPropDecimal(item, "iamt");
                decimal camt = GetPropDecimal(item, "camt");
                decimal samt = GetPropDecimal(item, "samt");
                decimal csamt = GetPropDecimal(item, "csamt");

                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                sheet.Cell(rowIdx, 3).SetValue("HSN");
                SetTextCell(sheet.Cell(rowIdx, 4), hsnSc);
                sheet.Cell(rowIdx, 5).SetValue(desc);
                sheet.Cell(rowIdx, 6).SetValue(uqc);
                SetNumberCell(sheet.Cell(rowIdx, 7), qty);
                SetNumberCell(sheet.Cell(rowIdx, 8), val);
                SetNumberCell(sheet.Cell(rowIdx, 9), txval);
                SetNumberCell(sheet.Cell(rowIdx, 10), rt);
                SetNumberCell(sheet.Cell(rowIdx, 11), iamt);
                SetNumberCell(sheet.Cell(rowIdx, 12), camt);
                SetNumberCell(sheet.Cell(rowIdx, 13), samt);
                SetNumberCell(sheet.Cell(rowIdx, 14), csamt);
                sheet.Cell(rowIdx, 15).SetValue("N");

                rowIdx++;
            }
        }

        private static void ProcessR1DocIssue(JsonElement root, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty("doc_issue", out var docIssueObj) || docIssueObj.ValueKind != JsonValueKind.Object) return;
            if (!docIssueObj.TryGetProperty("doc_det", out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var docGroup in arr.EnumerateArray())
            {
                int docNum = docGroup.TryGetProperty("doc_num", out var dn) ? dn.GetInt32() : 1;
                string docTypeDesc = docNum switch
                {
                    1 => "Invoices for outward supply",
                    2 => "Invoices for inward supply from unregistered person",
                    3 => "Revised Invoice",
                    4 => "Debit Note",
                    5 => "Credit Note",
                    _ => "Other Document"
                };

                if (docGroup.TryGetProperty("docs", out var docsArr) && docsArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var d in docsArr.EnumerateArray())
                    {
                        string from = GetPropString(d, "from");
                        string to = GetPropString(d, "to");
                        int totnum = d.TryGetProperty("totnum", out var tn) ? tn.GetInt32() : 0;
                        int canc = d.TryGetProperty("cancel", out var cn1) ? cn1.GetInt32() : (d.TryGetProperty("canc", out var cn2) ? cn2.GetInt32() : 0);

                        sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                        sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                        sheet.Cell(rowIdx, 3).SetValue(docTypeDesc);
                        SetTextCell(sheet.Cell(rowIdx, 4), from);
                        SetTextCell(sheet.Cell(rowIdx, 5), to);
                        sheet.Cell(rowIdx, 6).SetValue(totnum);
                        sheet.Cell(rowIdx, 7).SetValue(canc);
                        sheet.Cell(rowIdx, 8).SetValue("N");

                        rowIdx++;
                    }
                }
            }
        }

        #endregion

        #region R2A (GSTR-2A) Sheets

        private int BuildR2ASheets(XLWorkbook wb, List<(string FileName, JsonDocument Doc)> docs)
        {
            int totalRecords = 0;

            // Sheet: Purchase
            string[] purchCols = {
                "Company GSTIN", "Tax Period", "Doc Type", "Purchase Type", "Doc No", "Doc Date",
                "Supplier GSTIN", "Supplier Name", "Supplier State", "Place of Supply", "Port Code",
                "Reference Doc No", "Reference Doc Date", "Reverse Charge", "Doc Value",
                "Item Taxable Value", "GST Rate", "IGST", "CGST", "SGST", "Cess", "Is Amendment",
                "Original Doc No", "Original Doc Date", "GSTR-1 Status", "GSTR-1 Filing Date",
                "GSTR-3B Status", "Cancellation Date", "GSTR-9 (8A) ITC Available", "Uploaded By",
                "Source", "IRN Date", "IRN"
            };
            var purchSheet = wb.Worksheets.Add("Purchase");
            SetupTableHeaders(purchSheet, purchCols);
            int purchRowIdx = 2;

            // Sheet: ISD
            string[] isdCols = {
                "Company GSTIN", "Tax Period", "Doc Type", "Doc No", "Doc Date", "GSTIN", "Trade Name",
                "Reference Invoice No", "Reference Invoice Date", "IGST", "CGST", "SGST", "Cess",
                "ITC Eligible", "Is Amendment", "Original Doc No", "Original Doc Date", "Is Return Filed"
            };
            var isdSheet = wb.Worksheets.Add("ISD");
            SetupTableHeaders(isdSheet, isdCols);
            int isdRowIdx = 2;

            // Sheet: TDS
            string[] tdsCols = {
                "Company GSTIN", "Tax Period", "Deductor GSTIN", "Deductor Name", "Deductor Tax Period",
                "Taxable Value", "IGST", "CGST", "SGST"
            };
            var tdsSheet = wb.Worksheets.Add("TDS");
            SetupTableHeaders(tdsSheet, tdsCols);
            int tdsRowIdx = 2;

            // Sheet: TCS
            string[] tcsCols = {
                "Company GSTIN", "Tax Period", "Collector GSTIN", "Collector Name", "Collector Tax Period",
                "Gross Value of Supplies", "Value of Supplies Returned", "Net Amount Liable to TCS", "IGST", "CGST", "SGST"
            };
            var tcsSheet = wb.Worksheets.Add("TCS");
            SetupTableHeaders(tcsSheet, tcsCols);
            int tcsRowIdx = 2;

            foreach (var (fileName, doc) in docs)
            {
                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object) continue;

                var data = ResolveDataElement(root);

                string companyGstin = GetPropString(data, "gstin");
                if (string.IsNullOrEmpty(companyGstin)) companyGstin = GetPropString(root, "gstin");
                if (string.IsNullOrEmpty(companyGstin))
                {
                    var mGstin = Regex.Match(fileName, @"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}");
                    if (mGstin.Success) companyGstin = mGstin.Value;
                }
                if (string.IsNullOrEmpty(companyGstin)) companyGstin = "07ACWFS8659K2ZV";

                string rawFp = GetPropString(data, "fp", "ret_period", "rtnprd");
                if (string.IsNullOrEmpty(rawFp)) rawFp = GetPropString(root, "fp", "ret_period", "rtnprd");
                if (string.IsNullOrEmpty(rawFp))
                {
                    var mFp = Regex.Match(fileName, @"_([0-1][0-9]20[2-3][0-9])_");
                    if (mFp.Success) rawFp = mFp.Groups[1].Value;
                }
                string taxPeriod = FormatSinglePeriod(rawFp);

                JsonElement r2aData = data.TryGetProperty("b2b", out _) || data.TryGetProperty("cdn", out _) || data.TryGetProperty("tds", out _) ? data : root;

                // Process B2B / B2BA
                ProcessR2APurchaseArray(r2aData, "b2b", "B2B", "Invoice", "No", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);
                ProcessR2APurchaseArray(r2aData, "b2ba", "B2BA", "Invoice", "Yes", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);

                // Process CDN / CDNA
                ProcessR2ACdnArray(r2aData, "cdn", "CDNR", "No", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);
                ProcessR2ACdnArray(r2aData, "cdna", "CDNRA", "Yes", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);

                // Process ISD
                ProcessR2AIsd(r2aData, companyGstin, taxPeriod, isdSheet, ref isdRowIdx);

                // Process TDS
                ProcessR2ATds(r2aData, companyGstin, taxPeriod, tdsSheet, ref tdsRowIdx);

                // Process TCS
                ProcessR2ATcs(r2aData, companyGstin, taxPeriod, tcsSheet, ref tcsRowIdx);
            }

            FinalizeTableLayout(purchSheet, purchCols.Length, purchRowIdx);
            FinalizeTableLayout(isdSheet, isdCols.Length, isdRowIdx);
            FinalizeTableLayout(tdsSheet, tdsCols.Length, tdsRowIdx);
            FinalizeTableLayout(tcsSheet, tcsCols.Length, tcsRowIdx);

            totalRecords = (purchRowIdx - 2) + (isdRowIdx - 2) + (tdsRowIdx - 2) + (tcsRowIdx - 2);
            return totalRecords;
        }

        private static void ProcessR2APurchaseArray(
            JsonElement root, string propName, string purchaseType, string docType, string isAmendment,
            string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var supplier in arr.EnumerateArray())
            {
                string supplierGstin = GetPropString(supplier, "ctin");
                string supplierName = GetPropString(supplier, "cname", "trade_name");
                string supplierState = GstStateHelper.GetStateName(supplierGstin);
                string cfs = GetPropString(supplier, "cfs");
                string fldtr1 = GetPropString(supplier, "fldtr1");
                string cfs3b = GetPropString(supplier, "cfs3b");
                string dtcancel = GetPropString(supplier, "dtcancel");

                if (supplier.TryGetProperty("inv", out var invArr) && invArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var inv in invArr.EnumerateArray())
                    {
                        string docNo = GetPropString(inv, "inum");
                        string docDate = GetPropString(inv, "idt");
                        string pos = GetPropString(inv, "pos");
                        string rchrg = GetPropString(inv, "rchrg");
                        decimal docVal = GetPropDecimal(inv, "val");
                        string origDocNo = GetPropString(inv, "oinum");
                        string origDocDate = GetPropString(inv, "oidt");
                        string irn = GetPropString(inv, "irn");
                        string irnDate = GetPropString(inv, "irngendate", "irn_dt");

                        if (inv.TryGetProperty("itms", out var itmsArr) && itmsArr.ValueKind == JsonValueKind.Array && itmsArr.GetArrayLength() > 0)
                        {
                            foreach (var itm in itmsArr.EnumerateArray())
                            {
                                JsonElement det = itm.TryGetProperty("itm_det", out var d) ? d : itm;

                                decimal txval = GetPropDecimal(det, "txval");
                                decimal rt = GetPropDecimal(det, "rt");
                                decimal iamt = GetPropDecimal(det, "iamt");
                                decimal camt = GetPropDecimal(det, "camt");
                                decimal samt = GetPropDecimal(det, "samt");
                                decimal csamt = GetPropDecimal(det, "csamt");
                                if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                                sheet.Cell(rowIdx, 3).SetValue(docType);
                                sheet.Cell(rowIdx, 4).SetValue(purchaseType);
                                SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                                sheet.Cell(rowIdx, 6).SetValue(docDate);
                                SetTextCell(sheet.Cell(rowIdx, 7), supplierGstin);
                                sheet.Cell(rowIdx, 8).SetValue(supplierName);
                                sheet.Cell(rowIdx, 9).SetValue(supplierState);
                                sheet.Cell(rowIdx, 10).SetValue(FormatPos(pos));
                                sheet.Cell(rowIdx, 14).SetValue(rchrg);
                                SetNumberCell(sheet.Cell(rowIdx, 15), docVal);
                                SetNumberCell(sheet.Cell(rowIdx, 16), txval);
                                SetNumberCell(sheet.Cell(rowIdx, 17), rt);
                                SetNumberCell(sheet.Cell(rowIdx, 18), iamt);
                                SetNumberCell(sheet.Cell(rowIdx, 19), camt);
                                SetNumberCell(sheet.Cell(rowIdx, 20), samt);
                                SetNumberCell(sheet.Cell(rowIdx, 21), csamt);
                                sheet.Cell(rowIdx, 22).SetValue(isAmendment);
                                SetTextCell(sheet.Cell(rowIdx, 23), origDocNo);
                                sheet.Cell(rowIdx, 24).SetValue(origDocDate);
                                sheet.Cell(rowIdx, 25).SetValue(cfs == "Y" ? "Filed" : (cfs == "N" ? "Not Filed" : cfs));
                                sheet.Cell(rowIdx, 26).SetValue(fldtr1);
                                sheet.Cell(rowIdx, 27).SetValue(cfs3b == "Y" ? "Yes" : (cfs3b == "N" ? "No" : cfs3b));
                                sheet.Cell(rowIdx, 28).SetValue(dtcancel);
                                sheet.Cell(rowIdx, 29).SetValue("Yes");
                                sheet.Cell(rowIdx, 30).SetValue("Taxpayer");
                                sheet.Cell(rowIdx, 31).SetValue("GSTR-1");
                                sheet.Cell(rowIdx, 32).SetValue(irnDate);
                                SetTextCell(sheet.Cell(rowIdx, 33), irn);

                                rowIdx++;
                            }
                        }
                        else
                        {
                            decimal txval = GetPropDecimal(inv, "txval");
                            decimal rt = GetPropDecimal(inv, "rt");
                            decimal iamt = GetPropDecimal(inv, "iamt", "igst");
                            decimal camt = GetPropDecimal(inv, "camt", "cgst");
                            decimal samt = GetPropDecimal(inv, "samt", "sgst");
                            decimal csamt = GetPropDecimal(inv, "csamt", "cess");
                            if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                            sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                            sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                            sheet.Cell(rowIdx, 3).SetValue(docType);
                            sheet.Cell(rowIdx, 4).SetValue(purchaseType);
                            SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                            sheet.Cell(rowIdx, 6).SetValue(docDate);
                            SetTextCell(sheet.Cell(rowIdx, 7), supplierGstin);
                            sheet.Cell(rowIdx, 8).SetValue(supplierName);
                            sheet.Cell(rowIdx, 9).SetValue(supplierState);
                            sheet.Cell(rowIdx, 10).SetValue(FormatPos(pos));
                            sheet.Cell(rowIdx, 14).SetValue(rchrg);
                            SetNumberCell(sheet.Cell(rowIdx, 15), docVal);
                            SetNumberCell(sheet.Cell(rowIdx, 16), txval);
                            SetNumberCell(sheet.Cell(rowIdx, 17), rt);
                            SetNumberCell(sheet.Cell(rowIdx, 18), iamt);
                            SetNumberCell(sheet.Cell(rowIdx, 19), camt);
                            SetNumberCell(sheet.Cell(rowIdx, 20), samt);
                            SetNumberCell(sheet.Cell(rowIdx, 21), csamt);
                            sheet.Cell(rowIdx, 22).SetValue(isAmendment);
                            SetTextCell(sheet.Cell(rowIdx, 23), origDocNo);
                            sheet.Cell(rowIdx, 24).SetValue(origDocDate);
                            sheet.Cell(rowIdx, 25).SetValue(cfs == "Y" ? "Filed" : (cfs == "N" ? "Not Filed" : cfs));
                            sheet.Cell(rowIdx, 26).SetValue(fldtr1);
                            sheet.Cell(rowIdx, 27).SetValue(cfs3b == "Y" ? "Yes" : (cfs3b == "N" ? "No" : cfs3b));
                            sheet.Cell(rowIdx, 28).SetValue(dtcancel);
                            sheet.Cell(rowIdx, 29).SetValue("Yes");
                            sheet.Cell(rowIdx, 30).SetValue("Taxpayer");
                            sheet.Cell(rowIdx, 31).SetValue("GSTR-1");
                            sheet.Cell(rowIdx, 32).SetValue(irnDate);
                            SetTextCell(sheet.Cell(rowIdx, 33), irn);

                            rowIdx++;
                        }
                    }
                }
            }
        }

        private static void ProcessR2ACdnArray(
            JsonElement root, string propName, string purchaseType, string isAmendment,
            string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var supplier in arr.EnumerateArray())
            {
                string supplierGstin = GetPropString(supplier, "ctin");
                string supplierName = GetPropString(supplier, "cname", "trade_name");
                string supplierState = GstStateHelper.GetStateName(supplierGstin);
                string cfs = GetPropString(supplier, "cfs");
                string fldtr1 = GetPropString(supplier, "fldtr1");
                string cfs3b = GetPropString(supplier, "cfs3b");

                if (supplier.TryGetProperty("nt", out var ntArr) && ntArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var nt in ntArr.EnumerateArray())
                    {
                        string ntty = GetPropString(nt, "ntty");
                        string docType = ntty == "C" ? "Credit Note" : "Debit Note";
                        string docNo = GetPropString(nt, "nt_num", "inum");
                        string docDate = GetPropString(nt, "nt_dt", "idt");
                        string refDocNo = GetPropString(nt, "inum", "p_gst");
                        string refDocDate = GetPropString(nt, "idt");
                        decimal docVal = GetPropDecimal(nt, "val");
                        string pos = GetPropString(nt, "pos");
                        string rchrg = GetPropString(nt, "rchrg");

                        if (nt.TryGetProperty("itms", out var itmsArr) && itmsArr.ValueKind == JsonValueKind.Array && itmsArr.GetArrayLength() > 0)
                        {
                            foreach (var itm in itmsArr.EnumerateArray())
                            {
                                JsonElement det = itm.TryGetProperty("itm_det", out var d) ? d : itm;

                                decimal txval = GetPropDecimal(det, "txval");
                                decimal rt = GetPropDecimal(det, "rt");
                                decimal iamt = GetPropDecimal(det, "iamt");
                                decimal camt = GetPropDecimal(det, "camt");
                                decimal samt = GetPropDecimal(det, "samt");
                                decimal csamt = GetPropDecimal(det, "csamt");
                                if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                                sheet.Cell(rowIdx, 3).SetValue(docType);
                                sheet.Cell(rowIdx, 4).SetValue(purchaseType);
                                SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                                sheet.Cell(rowIdx, 6).SetValue(docDate);
                                SetTextCell(sheet.Cell(rowIdx, 7), supplierGstin);
                                sheet.Cell(rowIdx, 8).SetValue(supplierName);
                                sheet.Cell(rowIdx, 9).SetValue(supplierState);
                                sheet.Cell(rowIdx, 10).SetValue(FormatPos(pos));
                                SetTextCell(sheet.Cell(rowIdx, 12), refDocNo);
                                sheet.Cell(rowIdx, 13).SetValue(refDocDate);
                                sheet.Cell(rowIdx, 14).SetValue(rchrg);
                                SetNumberCell(sheet.Cell(rowIdx, 15), docVal);
                                SetNumberCell(sheet.Cell(rowIdx, 16), txval);
                                SetNumberCell(sheet.Cell(rowIdx, 17), rt);
                                SetNumberCell(sheet.Cell(rowIdx, 18), iamt);
                                SetNumberCell(sheet.Cell(rowIdx, 19), camt);
                                SetNumberCell(sheet.Cell(rowIdx, 20), samt);
                                SetNumberCell(sheet.Cell(rowIdx, 21), csamt);
                                sheet.Cell(rowIdx, 22).SetValue(isAmendment);
                                sheet.Cell(rowIdx, 25).SetValue(cfs == "Y" ? "Filed" : cfs);
                                sheet.Cell(rowIdx, 26).SetValue(fldtr1);
                                sheet.Cell(rowIdx, 27).SetValue(cfs3b == "Y" ? "Yes" : cfs3b);
                                sheet.Cell(rowIdx, 29).SetValue("Yes");

                                rowIdx++;
                            }
                        }
                        else
                        {
                            decimal txval = GetPropDecimal(nt, "txval");
                            decimal rt = GetPropDecimal(nt, "rt");
                            decimal iamt = GetPropDecimal(nt, "iamt", "igst");
                            decimal camt = GetPropDecimal(nt, "camt", "cgst");
                            decimal samt = GetPropDecimal(nt, "samt", "sgst");
                            decimal csamt = GetPropDecimal(nt, "csamt", "cess");
                            if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0) rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);

                            sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                            sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                            sheet.Cell(rowIdx, 3).SetValue(docType);
                            sheet.Cell(rowIdx, 4).SetValue(purchaseType);
                            SetTextCell(sheet.Cell(rowIdx, 5), docNo);
                            sheet.Cell(rowIdx, 6).SetValue(docDate);
                            SetTextCell(sheet.Cell(rowIdx, 7), supplierGstin);
                            sheet.Cell(rowIdx, 8).SetValue(supplierName);
                            sheet.Cell(rowIdx, 9).SetValue(supplierState);
                            sheet.Cell(rowIdx, 10).SetValue(FormatPos(pos));
                            SetTextCell(sheet.Cell(rowIdx, 12), refDocNo);
                            sheet.Cell(rowIdx, 13).SetValue(refDocDate);
                            sheet.Cell(rowIdx, 14).SetValue(rchrg);
                            SetNumberCell(sheet.Cell(rowIdx, 15), docVal);
                            SetNumberCell(sheet.Cell(rowIdx, 16), txval);
                            SetNumberCell(sheet.Cell(rowIdx, 17), rt);
                            SetNumberCell(sheet.Cell(rowIdx, 18), iamt);
                            SetNumberCell(sheet.Cell(rowIdx, 19), camt);
                            SetNumberCell(sheet.Cell(rowIdx, 20), samt);
                            SetNumberCell(sheet.Cell(rowIdx, 21), csamt);
                            sheet.Cell(rowIdx, 22).SetValue(isAmendment);
                            sheet.Cell(rowIdx, 25).SetValue(cfs == "Y" ? "Filed" : cfs);
                            sheet.Cell(rowIdx, 26).SetValue(fldtr1);
                            sheet.Cell(rowIdx, 27).SetValue(cfs3b == "Y" ? "Yes" : cfs3b);
                            sheet.Cell(rowIdx, 29).SetValue("Yes");

                            rowIdx++;
                        }
                    }
                }
            }
        }

        private static void ProcessR2AIsd(JsonElement root, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty("isd", out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var party in arr.EnumerateArray())
            {
                string isdGstin = GetPropString(party, "ctin");
                string tradeName = GetPropString(party, "cname", "trade_name");

                if (party.TryGetProperty("doclist", out var docList) && docList.ValueKind == JsonValueKind.Array)
                {
                    foreach (var docItem in docList.EnumerateArray())
                    {
                        string docType = GetPropString(docItem, "doctyp", "ty");
                        string docNo = GetPropString(docItem, "docnum", "num");
                        string docDate = GetPropString(docItem, "docdt", "dt");
                        decimal iamt = GetPropDecimal(docItem, "iamt");
                        decimal camt = GetPropDecimal(docItem, "camt");
                        decimal samt = GetPropDecimal(docItem, "samt");
                        decimal csamt = GetPropDecimal(docItem, "csamt");

                        sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                        sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                        sheet.Cell(rowIdx, 3).SetValue(docType);
                        SetTextCell(sheet.Cell(rowIdx, 4), docNo);
                        sheet.Cell(rowIdx, 5).SetValue(docDate);
                        SetTextCell(sheet.Cell(rowIdx, 6), isdGstin);
                        sheet.Cell(rowIdx, 7).SetValue(tradeName);
                        SetNumberCell(sheet.Cell(rowIdx, 10), iamt);
                        SetNumberCell(sheet.Cell(rowIdx, 11), camt);
                        SetNumberCell(sheet.Cell(rowIdx, 12), samt);
                        SetNumberCell(sheet.Cell(rowIdx, 13), csamt);
                        sheet.Cell(rowIdx, 14).SetValue("Y");
                        sheet.Cell(rowIdx, 15).SetValue("No");
                        sheet.Cell(rowIdx, 18).SetValue("Y");

                        rowIdx++;
                    }
                }
            }
        }

        private static void ProcessR2ATds(JsonElement root, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty("tds", out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var item in arr.EnumerateArray())
            {
                string deductorGstin = GetPropString(item, "gstin_deductor", "ctin");
                string deductorName = GetPropString(item, "deductor_name", "cname");
                decimal txval = GetPropDecimal(item, "amt_ded", "txval");
                string period = FormatSinglePeriod(GetPropString(item, "month", "fp"));
                if (string.IsNullOrEmpty(period)) period = taxPeriod;
                decimal iamt = GetPropDecimal(item, "iamt");
                decimal camt = GetPropDecimal(item, "camt");
                decimal samt = GetPropDecimal(item, "samt");

                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                SetTextCell(sheet.Cell(rowIdx, 3), deductorGstin);
                sheet.Cell(rowIdx, 4).SetValue(deductorName);
                sheet.Cell(rowIdx, 5).SetValue(period);
                SetNumberCell(sheet.Cell(rowIdx, 6), txval);
                SetNumberCell(sheet.Cell(rowIdx, 7), iamt);
                SetNumberCell(sheet.Cell(rowIdx, 8), camt);
                SetNumberCell(sheet.Cell(rowIdx, 9), samt);

                rowIdx++;
            }
        }

        private static void ProcessR2ATcs(JsonElement root, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!root.TryGetProperty("tcs", out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var item in arr.EnumerateArray())
            {
                string collectorGstin = GetPropString(item, "gstin_collector", "ctin");
                string collectorName = GetPropString(item, "collector_name", "cname");
                string period = FormatSinglePeriod(GetPropString(item, "month", "fp"));
                if (string.IsNullOrEmpty(period)) period = taxPeriod;
                decimal supVal = GetPropDecimal(item, "sup_val", "gross_val");
                decimal retVal = GetPropDecimal(item, "ret_val");
                decimal txval = GetPropDecimal(item, "txval", "net_val");
                decimal iamt = GetPropDecimal(item, "iamt");
                decimal camt = GetPropDecimal(item, "camt");
                decimal samt = GetPropDecimal(item, "samt");

                sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                SetTextCell(sheet.Cell(rowIdx, 3), collectorGstin);
                sheet.Cell(rowIdx, 4).SetValue(collectorName);
                sheet.Cell(rowIdx, 5).SetValue(period);
                SetNumberCell(sheet.Cell(rowIdx, 6), supVal);
                SetNumberCell(sheet.Cell(rowIdx, 7), retVal);
                SetNumberCell(sheet.Cell(rowIdx, 8), txval);
                SetNumberCell(sheet.Cell(rowIdx, 9), iamt);
                SetNumberCell(sheet.Cell(rowIdx, 10), camt);
                SetNumberCell(sheet.Cell(rowIdx, 11), samt);

                rowIdx++;
            }
        }


        #endregion

        #region R2B (GSTR-2B) Sheets

        private int BuildR2BSheets(XLWorkbook wb, List<(string FileName, JsonDocument Doc)> docs)
        {
            int totalRecords = 0;

            // Sheet: Purchase
            string[] purchCols = {
                "Company GSTIN", "Tax Period", "Doc Type", "Purchase Type", "Doc No", "Doc Date",
                "Supplier GSTIN", "Supplier Name", "Supplier State", "Place of Supply", "Port Code",
                "Reverse Charge", "Doc Value", "Item Taxable Value", "GST Rate", "IGST",
                "CGST", "SGST", "Cess", "Is Amendment", "Original Doc No", "Original Doc Date",
                "IMS Action", "GSTR-1 Filing Period", "GSTR-1 Filing Date", "ITC Eligible",
                "ITC Ineligible Reason Type", "GSTR-9 (8A) ITC Available", "ICEGATE Reference Date",
                "ICEGATE Received Date", "Uploaded By", "Source", "IRN Date", "IRN",
                "IMS ITC-IGST", "IMS ITC-CGST", "IMS ITC-SGST", "IMS ITC-Cess", "IMS Remarks"
            };
            var purchSheet = wb.Worksheets.Add("Purchase");
            SetupTableHeaders(purchSheet, purchCols);
            int purchRowIdx = 2;

            // Sheet: ISD
            string[] isdCols = {
                "Company GSTIN", "Tax Period", "Doc Type", "Doc No", "Doc Date", "GSTIN", "Trade Name",
                "Reference Invoice No", "Reference Invoice Date", "IGST", "CGST", "SGST", "Cess",
                "ITC Eligible", "Is Amendment", "Original Doc No", "Original Doc Date",
                "GSTR-6 Filing Period", "GSTR-6 Filing Date"
            };
            var isdSheet = wb.Worksheets.Add("ISD");
            SetupTableHeaders(isdSheet, isdCols);
            int isdRowIdx = 2;

            foreach (var (fileName, doc) in docs)
            {
                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object) continue;

                var data = ResolveDataElement(root);

                string companyGstin = GetPropString(data, "gstin");
                if (string.IsNullOrEmpty(companyGstin)) companyGstin = GetPropString(root, "gstin");
                if (string.IsNullOrEmpty(companyGstin))
                {
                    var mGstin = Regex.Match(fileName, @"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}");
                    if (mGstin.Success) companyGstin = mGstin.Value;
                }
                if (string.IsNullOrEmpty(companyGstin)) companyGstin = "07ACWFS8659K2ZV";

                string rawFp = GetPropString(data, "fp", "ret_period", "rtnprd");
                if (string.IsNullOrEmpty(rawFp)) rawFp = GetPropString(root, "fp", "ret_period", "rtnprd");
                if (string.IsNullOrEmpty(rawFp))
                {
                    var mFp = Regex.Match(fileName, @"_([0-1][0-9]20[2-3][0-9])_");
                    if (mFp.Success) rawFp = mFp.Groups[1].Value;
                }
                string taxPeriod = FormatSinglePeriod(rawFp);

                // In GSTR-2B, data is under data.docdata or root.docdata or direct sections
                JsonElement docData = data.TryGetProperty("docdata", out var dd) ? dd : (root.TryGetProperty("docdata", out var dd2) ? dd2 : data);

                // 1. Process B2B
                ProcessR2BPurchaseInvoices(docData, "b2b", "B2B", "Invoice", "No", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);

                // 2. Process B2BA
                ProcessR2BPurchaseInvoices(docData, "b2ba", "B2BA", "Invoice", "Yes", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);

                // 3. Process CDNR
                ProcessR2BPurchaseNotes(docData, "cdnr", "CDNR", "No", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);

                // 4. Process CDNRA
                ProcessR2BPurchaseNotes(docData, "cdnra", "CDNRA", "Yes", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);

                // 5. Process ECOM / ECO
                ProcessR2BPurchaseInvoices(docData, "ecom", "ECO", "Invoice", "No", companyGstin, taxPeriod, purchSheet, ref purchRowIdx);

                // 6. Process ISD
                ProcessR2BIsd(docData, companyGstin, taxPeriod, isdSheet, ref isdRowIdx);
            }

            FinalizeTableLayout(purchSheet, purchCols.Length, purchRowIdx);
            FinalizeTableLayout(isdSheet, isdCols.Length, isdRowIdx);

            totalRecords = (purchRowIdx - 2) + (isdRowIdx - 2);
            return totalRecords;
        }

        private static void ProcessR2BPurchaseInvoices(
            JsonElement docData, string propName, string purchaseType, string defaultDocType, string isAmendment,
            string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!docData.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var supplier in arr.EnumerateArray())
            {
                string supplierGstin = GetPropString(supplier, "ctin");
                string supplierName = GetPropString(supplier, "trdnm", "cname", "trade_name");
                string supplierState = GstStateHelper.GetStateName(supplierGstin);
                string supFildt = GetPropString(supplier, "supfildt", "fldt");
                string supPrd = FormatSinglePeriod(GetPropString(supplier, "supprd", "fp"));

                if (supplier.TryGetProperty("inv", out var invArr) && invArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var inv in invArr.EnumerateArray())
                    {
                        string docNo = GetPropString(inv, "inum", "nt_num");
                        string docDate = GetPropString(inv, "dt", "idt", "nt_dt");
                        string pos = GetPropString(inv, "pos");
                        string posFormatted = FormatPos(pos);
                        string rchrg = GetPropString(inv, "rev", "rchrg");
                        if (string.IsNullOrEmpty(rchrg)) rchrg = "N";
                        decimal docVal = GetPropDecimal(inv, "val");
                        string itcElg = GetPropString(inv, "itcavl");
                        string itcEligibleDesc = itcElg == "Y" ? "Yes" : (itcElg == "N" ? "No" : (string.IsNullOrEmpty(itcElg) ? "Yes" : itcElg));
                        string rsn = GetPropString(inv, "rsn");
                        string imsStatus = GetPropString(inv, "imsStatus", "ims_status", "imsAction");
                        string imsAction = imsStatus == "N" ? "No Action" : (imsStatus == "A" ? "Accepted" : (imsStatus == "R" ? "Rejected" : (imsStatus == "P" ? "Pending" : (string.IsNullOrEmpty(imsStatus) ? "No Action" : imsStatus))));
                        string irn = GetPropString(inv, "irn");
                        string irnDate = GetPropString(inv, "irngendate", "irn_dt");
                        string srctyp = GetPropString(inv, "srctyp");
                        if (string.IsNullOrEmpty(srctyp)) srctyp = !string.IsNullOrEmpty(irn) ? "E-Invoice" : "GSTR-1";
                        string oinum = GetPropString(inv, "oinum");
                        string oidt = GetPropString(inv, "oidt");
                        string invFildt = GetPropString(inv, "fldt", "supfildt");
                        if (string.IsNullOrEmpty(invFildt)) invFildt = supFildt;
                        string invPrd = FormatSinglePeriod(GetPropString(inv, "fp", "supprd"));
                        if (string.IsNullOrEmpty(invPrd)) invPrd = supPrd;

                        if (inv.TryGetProperty("itms", out var itmsArr) && itmsArr.ValueKind == JsonValueKind.Array && itmsArr.GetArrayLength() > 0)
                        {
                            foreach (var itm in itmsArr.EnumerateArray())
                            {
                                JsonElement det = itm.TryGetProperty("itm_det", out var d) ? d : itm;

                                decimal txval = GetPropDecimal(det, "txval");
                                decimal rt = GetPropDecimal(det, "rt");
                                decimal iamt = GetPropDecimal(det, "iamt", "igst");
                                decimal camt = GetPropDecimal(det, "camt", "cgst");
                                decimal samt = GetPropDecimal(det, "samt", "sgst");
                                decimal csamt = GetPropDecimal(det, "csamt", "cess");
                                if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0)
                                {
                                    rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);
                                }

                                WritePurchaseRow(sheet, rowIdx++, companyGstin, taxPeriod, defaultDocType, purchaseType,
                                    docNo, docDate, supplierGstin, supplierName, supplierState, posFormatted,
                                    rchrg, docVal, txval, rt, iamt, camt, samt, csamt,
                                    isAmendment, oinum, oidt, imsAction, invPrd, invFildt,
                                    itcEligibleDesc, rsn, srctyp, irnDate, irn);
                            }
                        }
                        else
                        {
                            decimal txval = GetPropDecimal(inv, "txval");
                            decimal rt = GetPropDecimal(inv, "rt");
                            decimal iamt = GetPropDecimal(inv, "iamt", "igst");
                            decimal camt = GetPropDecimal(inv, "camt", "cgst");
                            decimal samt = GetPropDecimal(inv, "samt", "sgst");
                            decimal csamt = GetPropDecimal(inv, "csamt", "cess");
                            if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0)
                            {
                                rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);
                            }

                            WritePurchaseRow(sheet, rowIdx++, companyGstin, taxPeriod, defaultDocType, purchaseType,
                                docNo, docDate, supplierGstin, supplierName, supplierState, posFormatted,
                                rchrg, docVal, txval, rt, iamt, camt, samt, csamt,
                                isAmendment, oinum, oidt, imsAction, invPrd, invFildt,
                                itcEligibleDesc, rsn, srctyp, irnDate, irn);
                        }
                    }
                }
            }
        }

        private static void ProcessR2BPurchaseNotes(
            JsonElement docData, string propName, string purchaseType, string isAmendment,
            string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!docData.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var supplier in arr.EnumerateArray())
            {
                string supplierGstin = GetPropString(supplier, "ctin");
                string supplierName = GetPropString(supplier, "trdnm", "cname", "trade_name");
                string supplierState = GstStateHelper.GetStateName(supplierGstin);
                string supFildt = GetPropString(supplier, "supfildt", "fldt");
                string supPrd = FormatSinglePeriod(GetPropString(supplier, "supprd", "fp"));

                JsonElement notesArr = default;
                if (supplier.TryGetProperty("nt", out var nt) && nt.ValueKind == JsonValueKind.Array) notesArr = nt;
                else if (supplier.TryGetProperty("inv", out var inv) && inv.ValueKind == JsonValueKind.Array) notesArr = inv;

                if (notesArr.ValueKind != JsonValueKind.Array) continue;

                foreach (var note in notesArr.EnumerateArray())
                {
                    string typ = GetPropString(note, "typ", "ntty");
                    string docType = (typ == "C" || typ.Equals("Credit", StringComparison.OrdinalIgnoreCase)) ? "Credit Note"
                                   : ((typ == "D" || typ.Equals("Debit", StringComparison.OrdinalIgnoreCase)) ? "Debit Note" : "Credit Note");

                    string docNo = GetPropString(note, "ntnum", "nt_num", "inum");
                    string docDate = GetPropString(note, "dt", "idt", "nt_dt");
                    string pos = GetPropString(note, "pos");
                    string posFormatted = FormatPos(pos);
                    string rchrg = GetPropString(note, "rev", "rchrg");
                    if (string.IsNullOrEmpty(rchrg)) rchrg = "N";
                    decimal docVal = GetPropDecimal(note, "val");
                    string itcElg = GetPropString(note, "itcavl");
                    string itcEligibleDesc = itcElg == "Y" ? "Yes" : (itcElg == "N" ? "No" : (string.IsNullOrEmpty(itcElg) ? "Yes" : itcElg));
                    string rsn = GetPropString(note, "rsn");
                    string imsStatus = GetPropString(note, "imsStatus", "ims_status", "imsAction");
                    string imsAction = imsStatus == "N" ? "No Action" : (imsStatus == "A" ? "Accepted" : (imsStatus == "R" ? "Rejected" : (imsStatus == "P" ? "Pending" : (string.IsNullOrEmpty(imsStatus) ? "No Action" : imsStatus))));
                    string irn = GetPropString(note, "irn");
                    string irnDate = GetPropString(note, "irngendate", "irn_dt");
                    string srctyp = GetPropString(note, "srctyp");
                    if (string.IsNullOrEmpty(srctyp)) srctyp = !string.IsNullOrEmpty(irn) ? "E-Invoice" : "GSTR-1";
                    string oinum = GetPropString(note, "ont_num", "oinum");
                    string oidt = GetPropString(note, "ont_dt", "oidt");
                    string noteFildt = GetPropString(note, "fldt", "supfildt");
                    if (string.IsNullOrEmpty(noteFildt)) noteFildt = supFildt;
                    string notePrd = FormatSinglePeriod(GetPropString(note, "fp", "supprd"));
                    if (string.IsNullOrEmpty(notePrd)) notePrd = supPrd;

                    if (note.TryGetProperty("itms", out var itmsArr) && itmsArr.ValueKind == JsonValueKind.Array && itmsArr.GetArrayLength() > 0)
                    {
                        foreach (var itm in itmsArr.EnumerateArray())
                        {
                            JsonElement det = itm.TryGetProperty("itm_det", out var d) ? d : itm;

                            decimal txval = GetPropDecimal(det, "txval");
                            decimal rt = GetPropDecimal(det, "rt");
                            decimal iamt = GetPropDecimal(det, "iamt", "igst");
                            decimal camt = GetPropDecimal(det, "camt", "cgst");
                            decimal samt = GetPropDecimal(det, "samt", "sgst");
                            decimal csamt = GetPropDecimal(det, "csamt", "cess");
                            if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0)
                            {
                                rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);
                            }

                            WritePurchaseRow(sheet, rowIdx++, companyGstin, taxPeriod, docType, purchaseType,
                                docNo, docDate, supplierGstin, supplierName, supplierState, posFormatted,
                                rchrg, docVal, txval, rt, iamt, camt, samt, csamt,
                                isAmendment, oinum, oidt, imsAction, notePrd, noteFildt,
                                itcEligibleDesc, rsn, srctyp, irnDate, irn);
                        }
                    }
                    else
                    {
                        decimal txval = GetPropDecimal(note, "txval");
                        decimal rt = GetPropDecimal(note, "rt");
                        decimal iamt = GetPropDecimal(note, "iamt", "igst");
                        decimal camt = GetPropDecimal(note, "camt", "cgst");
                        decimal samt = GetPropDecimal(note, "samt", "sgst");
                        decimal csamt = GetPropDecimal(note, "csamt", "cess");
                        if (rt == 0 && txval > 0 && (iamt + camt + samt) > 0)
                        {
                            rt = Math.Round((iamt + camt + samt) * 100m / txval, 2);
                        }

                        WritePurchaseRow(sheet, rowIdx++, companyGstin, taxPeriod, docType, purchaseType,
                            docNo, docDate, supplierGstin, supplierName, supplierState, posFormatted,
                            rchrg, docVal, txval, rt, iamt, camt, samt, csamt,
                            isAmendment, oinum, oidt, imsAction, notePrd, noteFildt,
                            itcEligibleDesc, rsn, srctyp, irnDate, irn);
                    }
                }
            }
        }

        private static void WritePurchaseRow(
            IXLWorksheet sheet, int row, string companyGstin, string taxPeriod, string docType, string purchaseType,
            string docNo, string docDate, string supplierGstin, string supplierName, string supplierState, string pos,
            string rchrg, decimal docVal, decimal txval, decimal rt, decimal iamt, decimal camt, decimal samt, decimal csamt,
            string isAmendment, string origDocNo, string origDocDate, string imsAction, string gstr1Fp, string gstr1Fldt,
            string itcEligible, string itcReason, string source, string irnDate, string irn)
        {
            sheet.Cell(row, 1).SetValue(companyGstin);
            sheet.Cell(row, 2).SetValue(taxPeriod);
            sheet.Cell(row, 3).SetValue(docType);
            sheet.Cell(row, 4).SetValue(purchaseType);
            SetTextCell(sheet.Cell(row, 5), docNo);
            sheet.Cell(row, 6).SetValue(docDate);
            SetTextCell(sheet.Cell(row, 7), supplierGstin);
            sheet.Cell(row, 8).SetValue(supplierName);
            sheet.Cell(row, 9).SetValue(supplierState);
            sheet.Cell(row, 10).SetValue(pos);
            sheet.Cell(row, 11).SetValue(string.Empty); // Port Code
            sheet.Cell(row, 12).SetValue(rchrg);
            SetNumberCell(sheet.Cell(row, 13), docVal);
            SetNumberCell(sheet.Cell(row, 14), txval);
            SetNumberCell(sheet.Cell(row, 15), rt);
            SetNumberCell(sheet.Cell(row, 16), iamt);
            SetNumberCell(sheet.Cell(row, 17), camt);
            SetNumberCell(sheet.Cell(row, 18), samt);
            SetNumberCell(sheet.Cell(row, 19), csamt);
            sheet.Cell(row, 20).SetValue(isAmendment);
            SetTextCell(sheet.Cell(row, 21), origDocNo);
            sheet.Cell(row, 22).SetValue(origDocDate);
            sheet.Cell(row, 23).SetValue(imsAction);
            sheet.Cell(row, 24).SetValue(gstr1Fp);
            sheet.Cell(row, 25).SetValue(gstr1Fldt);
            sheet.Cell(row, 26).SetValue(itcEligible);
            sheet.Cell(row, 27).SetValue(itcReason);
            sheet.Cell(row, 28).SetValue("Yes"); // GSTR-9 (8A) ITC Available
            sheet.Cell(row, 29).SetValue(string.Empty); // ICEGATE Reference Date
            sheet.Cell(row, 30).SetValue(string.Empty); // ICEGATE Received Date
            sheet.Cell(row, 31).SetValue("Taxpayer"); // Uploaded By
            sheet.Cell(row, 32).SetValue(source); // Source
            sheet.Cell(row, 33).SetValue(irnDate);
            SetTextCell(sheet.Cell(row, 34), irn);
            sheet.Cell(row, 35).SetValue(string.Empty); // IMS ITC-IGST
            sheet.Cell(row, 36).SetValue(string.Empty); // IMS ITC-CGST
            sheet.Cell(row, 37).SetValue(string.Empty); // IMS ITC-SGST
            sheet.Cell(row, 38).SetValue(string.Empty); // IMS ITC-Cess
            sheet.Cell(row, 39).SetValue(string.Empty); // IMS Remarks
        }

        private static void ProcessR2BIsd(JsonElement docData, string companyGstin, string taxPeriod, IXLWorksheet sheet, ref int rowIdx)
        {
            if (!docData.TryGetProperty("isd", out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var party in arr.EnumerateArray())
            {
                string isdGstin = GetPropString(party, "ctin");
                string tradeName = GetPropString(party, "cname", "trade_name");

                if (party.TryGetProperty("doclist", out var docList) && docList.ValueKind == JsonValueKind.Array)
                {
                    foreach (var docItem in docList.EnumerateArray())
                    {
                        string docType = GetPropString(docItem, "doctyp", "ty");
                        string docNo = GetPropString(docItem, "docnum", "num");
                        string docDate = GetPropString(docItem, "docdt", "dt");
                        decimal iamt = GetPropDecimal(docItem, "iamt");
                        decimal camt = GetPropDecimal(docItem, "camt");
                        decimal samt = GetPropDecimal(docItem, "samt");
                        decimal csamt = GetPropDecimal(docItem, "csamt");
                        string itcavl = GetPropString(docItem, "itcavl");

                        sheet.Cell(rowIdx, 1).SetValue(companyGstin);
                        sheet.Cell(rowIdx, 2).SetValue(taxPeriod);
                        sheet.Cell(rowIdx, 3).SetValue(docType);
                        SetTextCell(sheet.Cell(rowIdx, 4), docNo);
                        sheet.Cell(rowIdx, 5).SetValue(docDate);
                        SetTextCell(sheet.Cell(rowIdx, 6), isdGstin);
                        sheet.Cell(rowIdx, 7).SetValue(tradeName);
                        SetNumberCell(sheet.Cell(rowIdx, 10), iamt);
                        SetNumberCell(sheet.Cell(rowIdx, 11), camt);
                        SetNumberCell(sheet.Cell(rowIdx, 12), samt);
                        SetNumberCell(sheet.Cell(rowIdx, 13), csamt);
                        sheet.Cell(rowIdx, 14).SetValue(itcavl == "Y" ? "Yes" : (itcavl == "N" ? "No" : itcavl));
                        sheet.Cell(rowIdx, 15).SetValue("No");

                        rowIdx++;
                    }
                }
            }
        }

        #endregion

        #region Audit Index & Styling Helpers

        private static void BuildAuditIndexSheet(IXLWorksheet sheet, Dictionary<string, JsonFlatValue> leaves)
        {
            string[] headers = { "#", "JSON Path", "Field Name", "Value", "Data Type" };
            SetupTableHeaders(sheet, headers);

            int rowIdx = 2;
            int id = 1;
            foreach (var kvp in leaves)
            {
                sheet.Cell(rowIdx, 1).SetValue(id);
                SetTextCell(sheet.Cell(rowIdx, 2), kvp.Key);
                SetTextCell(sheet.Cell(rowIdx, 3), ExtractFieldName(kvp.Key));
                SetTextCell(sheet.Cell(rowIdx, 4), kvp.Value.RawString);
                sheet.Cell(rowIdx, 5).SetValue(kvp.Value.ValueKind.ToString());
                rowIdx++;
                id++;
            }

            FinalizeTableLayout(sheet, headers.Length, rowIdx);
        }

        private static void SetupTableHeaders(IXLWorksheet sheet, string[] columnHeaders)
        {
            for (int i = 0; i < columnHeaders.Length; i++)
            {
                var cell = sheet.Cell(1, i + 1);
                cell.SetValue(columnHeaders[i]);
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontSize = 10;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = RoyalBlue;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            sheet.Row(1).Height = 26;
            sheet.SheetView.FreezeRows(1);
        }

        private static void FinalizeTableLayout(IXLWorksheet sheet, int colCount, int totalRows)
        {
            if (totalRows > 2 && colCount > 0)
            {
                sheet.Range(1, 1, totalRows - 1, colCount).SetAutoFilter();
            }

            try
            {
                sheet.Columns(1, Math.Min(colCount, 50)).AdjustToContents(1, Math.Min(totalRows, 300), 10.0, 50.0);
            }
            catch
            {
                for (int c = 1; c <= colCount; c++) sheet.Column(c).Width = 16;
            }
        }

        private static void SetTextCell(IXLCell cell, string value)
        {
            cell.Style.NumberFormat.Format = "@";
            cell.SetValue(value ?? string.Empty);
        }

        private static void SetNumberCell(IXLCell cell, decimal value)
        {
            cell.Style.NumberFormat.Format = "#,##0.00";
            cell.SetValue(value);
        }

        private static JsonElement ResolveDataElement(JsonElement root)
        {
            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("data", out var d) && d.ValueKind == JsonValueKind.Object)
            {
                return d;
            }
            return root;
        }

        private static string GetPropString(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var prop))
                {
                    if (prop.ValueKind == JsonValueKind.String) return prop.GetString() ?? string.Empty;
                    return prop.GetRawText().Trim('"');
                }
            }
            return string.Empty;
        }

        private static decimal GetPropDecimal(JsonElement element, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (element.TryGetProperty(name, out var prop))
                {
                    if (prop.ValueKind == JsonValueKind.Number && prop.TryGetDecimal(out decimal d)) return d;
                    string raw = prop.GetRawText().Trim('"');
                    if (decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsed)) return parsed;
                }
            }
            return 0m;
        }

        private static string FormatSinglePeriod(string? fp)
        {
            if (string.IsNullOrWhiteSpace(fp)) return string.Empty;
            fp = fp.Trim();
            if (fp.Length == 6 && int.TryParse(fp[..2], out int month) && int.TryParse(fp[2..], out int year) && month >= 1 && month <= 12)
            {
                var dt = new DateTime(year, month, 1);
                return dt.ToString("MMM yyyy");
            }
            return fp;
        }

        private static string FormatPeriodRange(HashSet<string> fps)
        {
            if (fps.Count == 0) return "Apr 2025-Mar 2026";
            if (fps.Count == 1) return FormatSinglePeriod(fps.First());

            var dates = new List<DateTime>();
            foreach (var fp in fps)
            {
                if (fp.Length == 6 && int.TryParse(fp[..2], out int month) && int.TryParse(fp[2..], out int year) && month >= 1 && month <= 12)
                {
                    dates.Add(new DateTime(year, month, 1));
                }
            }

            if (dates.Count > 0)
            {
                dates.Sort();
                var min = dates.First();
                var max = dates.Last();
                if (min == max) return min.ToString("MMM yyyy");
                return $"{min:MMM yyyy}-{max:MMM yyyy}";
            }

            return "Apr 2025-Mar 2026";
        }

        private static string FormatPos(string? pos)
        {
            if (string.IsNullOrWhiteSpace(pos)) return string.Empty;
            pos = pos.Trim();
            if (pos.Length <= 2 && int.TryParse(pos, out _))
            {
                string state = GstStateHelper.GetStateName(pos.PadLeft(2, '0'));
                return !string.IsNullOrEmpty(state) ? $"{pos.PadLeft(2, '0')}-{state}" : pos;
            }
            return pos;
        }

        private static void IndexLeaves(JsonElement element, string currentPath, Dictionary<string, JsonFlatValue> leaves)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var prop in element.EnumerateObject())
                    {
                        IndexLeaves(prop.Value, $"{currentPath}.{prop.Name}", leaves);
                    }
                    break;
                case JsonValueKind.Array:
                    int idx = 0;
                    foreach (var item in element.EnumerateArray())
                    {
                        IndexLeaves(item, $"{currentPath}[{idx++}]", leaves);
                    }
                    break;
                default:
                    string fieldName = ExtractFieldName(currentPath);
                    leaves[currentPath] = JsonFlatValue.FromJsonElement(element, fieldName);
                    break;
            }
        }

        private static string ExtractFieldName(string path)
        {
            int dot = path.LastIndexOf('.');
            string candidate = dot >= 0 ? path[(dot + 1)..] : path;
            int bracket = candidate.IndexOf('[');
            return bracket >= 0 ? candidate[..bracket] : candidate;
        }

        #endregion
    }
}
