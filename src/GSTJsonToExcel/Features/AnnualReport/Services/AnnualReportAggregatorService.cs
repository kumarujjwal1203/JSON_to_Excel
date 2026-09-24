using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using GSTJsonToExcel.Features.AnnualReport.Models;
using GSTJsonToExcel.Helpers;

namespace GSTJsonToExcel.Features.AnnualReport.Services
{
    public class AnnualReportAggregatorService
    {
        public AnnualWorkbookData BuildAnnualData(
            List<AnnualScannedFile> allValidFiles,
            AnnualReportUserSettings settings,
            string targetFy)
        {
            ParseFyYears(targetFy, allValidFiles, out int startYear, out int endYear);
            var fyStartDate = new DateTime(startYear, 4, 1);
            string resolvedFy = $"{startYear}-{(endYear % 100):D2}";

            var gstins = allValidFiles
                .Select(f => f.Gstin)
                .Where(g => !string.IsNullOrWhiteSpace(g))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var displayMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var gstin in gstins)
            {
                displayMap[gstin] = ResolveGstinDisplay(gstin, settings);
            }

            var workbookData = new AnnualWorkbookData
            {
                CompanyName = string.IsNullOrWhiteSpace(settings.CompanyName)
                    ? (gstins.Count > 0 ? $"GSTIN_{gstins[0]}" : "GST_Taxpayer")
                    : settings.CompanyName.Trim(),
                FinancialYear = resolvedFy,
                PeriodRangeLabel = $"Apr {startYear} to Mar {endYear}",
                StartYear = startYear,
                EndYear = endYear,
                Gstins = gstins,
                GstinDisplayMap = displayMap,
                Has3BData = allValidFiles.Any(f => f.ReturnType == AnnualReturnType.GSTR3B),
                Has1Data = allValidFiles.Any(f => f.ReturnType == AnnualReturnType.GSTR1),
                Has2AData = allValidFiles.Any(f => f.ReturnType == AnnualReturnType.GSTR2A),
                Has2BData = allValidFiles.Any(f => f.ReturnType == AnnualReturnType.GSTR2B)
            };

            foreach (var gstin in gstins)
            {
                string gstinDisplay = displayMap[gstin];
                var gstinFiles = allValidFiles
                    .Where(f => string.Equals(f.Gstin, gstin, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                var r3BFiles = gstinFiles.Where(f => f.ReturnType == AnnualReturnType.GSTR3B).ToList();
                var r1Files = gstinFiles.Where(f => f.ReturnType == AnnualReturnType.GSTR1).ToList();
                var r2AFiles = gstinFiles.Where(f => f.ReturnType == AnnualReturnType.GSTR2A).ToList();
                var r2BFiles = gstinFiles.Where(f => f.ReturnType == AnnualReturnType.GSTR2B).ToList();

                // GSTR-3B has 12 month columns (Apr startYear .. Mar endYear)
                var gstin3BRows = BuildGstr3BRowsForGstin(gstin, gstinDisplay, r3BFiles, startYear);

                // GSTR-1, GSTR-2A, GSTR-2B have 24 month columns (Apr startYear .. Mar endYear + Apr endYear .. Mar endYear+1)
                var gstin1Rows = BuildGstr1RowsForGstin(gstin, gstinDisplay, r1Files, startYear, fyStartDate);
                var gstin2ARows = BuildGstr2ARowsForGstin(gstin, gstinDisplay, r2AFiles, startYear, fyStartDate);
                var gstin2BRows = BuildGstr2BRowsForGstin(gstin, gstinDisplay, r2BFiles, startYear, fyStartDate);

                workbookData.Gstr3BRows.AddRange(gstin3BRows);
                workbookData.Gstr1Rows.AddRange(gstin1Rows);
                workbookData.Gstr2ARows.AddRange(gstin2ARows);
                workbookData.Gstr2BRows.AddRange(gstin2BRows);

                // Build Overview Summary for this GSTIN
                var overview = BuildOverviewSummaryForGstin(
                    gstin, gstinDisplay,
                    r3BFiles, r1Files, r2AFiles, r2BFiles,
                    gstin3BRows, gstin1Rows, gstin2ARows, gstin2BRows);
                workbookData.OverviewSummaries.Add(overview);

                // Build Reconciliation Sheets (12 month columns: Apr startYear .. Mar endYear)
                if (workbookData.Has3BData && workbookData.Has1Data)
                {
                    var vs1 = Build3BVs1ForGstin(gstin, gstinDisplay, gstin3BRows, gstin1Rows, workbookData.ReconciliationAlerts);
                    workbookData.Gstr3BVs1Rows.AddRange(vs1);
                }

                if (workbookData.Has3BData && workbookData.Has2AData)
                {
                    var vs2A = Build3BVs2AOr2BForGstin(gstin, gstinDisplay, "GSTR-2A", gstin3BRows, gstin2ARows, workbookData.ReconciliationAlerts);
                    workbookData.Gstr3BVs2ARows.AddRange(vs2A);
                }

                if (workbookData.Has3BData && workbookData.Has2BData)
                {
                    var vs2B = Build3BVs2AOr2BForGstin(gstin, gstinDisplay, "GSTR-2B", gstin3BRows, gstin2BRows, workbookData.ReconciliationAlerts);
                    workbookData.Gstr3BVs2BRows.AddRange(vs2B);
                }
            }

            return workbookData;
        }

        /// <summary>
        /// Maps a file's (YearNumber, MonthNumber) to a 0..23 index relative to Apr {startYear}.
        /// Returns 0..11 for Apr startYear..Mar (startYear+1), and 12..23 for Apr (startYear+1)..Mar (startYear+2).
        /// </summary>
        private static int Get24MonthSlot(AnnualScannedFile file, int startYear)
        {
            int monthOffset = ((file.YearNumber - startYear) * 12) + (file.MonthNumber - 4);
            if (monthOffset >= 0 && monthOffset < 24)
            {
                return monthOffset;
            }
            // Fallback to FyMonthIndex (0..11) if user uploaded a single month from another year
            return file.FyMonthIndex >= 0 && file.FyMonthIndex < 12 ? file.FyMonthIndex : -1;
        }

        private static int Get12MonthSlot(AnnualScannedFile file, int startYear)
        {
            int monthOffset = ((file.YearNumber - startYear) * 12) + (file.MonthNumber - 4);
            if (monthOffset >= 0 && monthOffset < 12)
            {
                return monthOffset;
            }
            return file.FyMonthIndex >= 0 && file.FyMonthIndex < 12 ? file.FyMonthIndex : -1;
        }

        public static string ResolveGstinDisplay(string gstin, AnnualReportUserSettings settings)
        {
            string state = GstStateHelper.GetStateName(gstin);
            return settings.GstinDisplayMode switch
            {
                GstinDisplayMode.ShowStateOnly => !string.IsNullOrWhiteSpace(state) ? state : gstin,
                GstinDisplayMode.ShowCustomBranchName =>
                    settings.CustomBranchLabels != null &&
                    settings.CustomBranchLabels.TryGetValue(gstin, out var custom) &&
                    !string.IsNullOrWhiteSpace(custom)
                        ? custom.Trim()
                        : GstStateHelper.FormatGstinWithState(gstin),
                _ => GstStateHelper.FormatGstinWithState(gstin)
            };
        }

        private static void ParseFyYears(string targetFy, List<AnnualScannedFile> files, out int startYear, out int endYear)
        {
            startYear = DateTime.Now.Year;
            if (!string.IsNullOrWhiteSpace(targetFy) && targetFy.Length >= 4 && int.TryParse(targetFy[..4], out int parsed))
            {
                startYear = parsed;
            }
            else
            {
                var dominantFy = files
                    .Where(f => !string.IsNullOrWhiteSpace(f.FinancialYear) && f.FinancialYear.Length >= 4)
                    .GroupBy(f => f.FinancialYear)
                    .OrderByDescending(g => g.Count())
                    .FirstOrDefault()?.Key;

                if (dominantFy != null && int.TryParse(dominantFy[..4], out int fyParsed))
                {
                    startYear = fyParsed;
                }
            }
            endYear = startYear + 1;
        }

        #region 1. GSTR-3B Builder (12 Monthly Columns: Apr YYYY .. Mar YYYY+1 + Total)

        private List<AnnualSheetRow> BuildGstr3BRowsForGstin(
            string gstin,
            string gstinDisplay,
            List<AnnualScannedFile> files,
            int startYear)
        {
            var rows = new List<AnnualSheetRow>();
            var map = new Dictionary<string, AnnualSheetRow>(StringComparer.OrdinalIgnoreCase);

            AnnualSheetRow AddRow(string section, string typeName, bool isSummary = false)
            {
                var r = new AnnualSheetRow(12)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = section,
                    Description = typeName,
                    SortOrder = rows.Count + 1,
                    IsSectionSummaryRow = isSummary
                };
                rows.Add(r);
                map[$"{section}|{typeName}"] = r;
                return r;
            }

            // 3.1.A
            const string s31A = "3.1.A Outward taxable supplies (excluding zero rated)";
            AddRow(s31A, "Taxable");
            AddRow(s31A, "IGST");
            AddRow(s31A, "CGST");
            AddRow(s31A, "SGST");
            AddRow(s31A, "Cess");

            // 3.1.B
            const string s31B = "3.1.B Outward taxable supplies (zero rated)";
            AddRow(s31B, "Taxable");
            AddRow(s31B, "IGST");
            AddRow(s31B, "Cess");

            // 3.1.C
            const string s31C = "3.1.C Other outward supplies (nil-rated, exempted)";
            AddRow(s31C, "Taxable");

            // 3.1.D
            const string s31D = "3.1.D Inward taxable supplies (reverse charge)";
            AddRow(s31D, "Taxable");
            AddRow(s31D, "IGST");
            AddRow(s31D, "CGST");
            AddRow(s31D, "SGST");
            AddRow(s31D, "Cess");

            // 3.1.E
            const string s31E = "3.1.E Non-GST outward supplies";
            AddRow(s31E, "Taxable");

            // 3.1.1 ECO
            const string s311_i = "3.1.1(i) Supplies on which ECO pays tax u/s 9(5)";
            AddRow(s311_i, "Taxable");
            AddRow(s311_i, "IGST");
            AddRow(s311_i, "CGST");
            AddRow(s311_i, "SGST");
            AddRow(s311_i, "Cess");

            const string s311_ii = "3.1.1(ii) Supplies made through ECO";
            AddRow(s311_ii, "Taxable");

            // 4.A Eligible ITC
            const string s4A1 = "4.A.1 Import of goods";
            AddRow(s4A1, "IGST");
            AddRow(s4A1, "Cess");

            const string s4A2 = "4.A.2 Import of services";
            AddRow(s4A2, "IGST");
            AddRow(s4A2, "Cess");

            const string s4A3 = "4.A.3 Inward supplies liable to reverse charge";
            AddRow(s4A3, "IGST");
            AddRow(s4A3, "CGST");
            AddRow(s4A3, "SGST");
            AddRow(s4A3, "Cess");

            const string s4A4 = "4.A.4 Inward supplies from ISD";
            AddRow(s4A4, "IGST");
            AddRow(s4A4, "CGST");
            AddRow(s4A4, "SGST");
            AddRow(s4A4, "Cess");

            const string s4A5 = "4.A.5 All other ITC";
            AddRow(s4A5, "IGST");
            AddRow(s4A5, "CGST");
            AddRow(s4A5, "SGST");
            AddRow(s4A5, "Cess");

            // 4.B ITC Reversed
            const string s4B1 = "4.B.1 ITC Reversed - As per rules 38, 42 & 43 and sec 17(5)";
            AddRow(s4B1, "IGST");
            AddRow(s4B1, "CGST");
            AddRow(s4B1, "SGST");
            AddRow(s4B1, "Cess");

            const string s4B2 = "4.B.2 ITC Reversed - Others";
            AddRow(s4B2, "IGST");
            AddRow(s4B2, "CGST");
            AddRow(s4B2, "SGST");
            AddRow(s4B2, "Cess");

            // Net ITC (Section Summary Row -> #BDD7EE fill + bold)
            const string sNetItc = "4(C) Net ITC Available (4A - 4B)";
            AddRow(sNetItc, "IGST", isSummary: true);
            AddRow(sNetItc, "CGST", isSummary: true);
            AddRow(sNetItc, "SGST", isSummary: true);
            AddRow(sNetItc, "Cess", isSummary: true);

            // 4.D Ineligible ITC
            const string s4D1 = "4.D.1 Ineligible ITC - As per section 17(5)";
            AddRow(s4D1, "IGST");
            AddRow(s4D1, "CGST");
            AddRow(s4D1, "SGST");
            AddRow(s4D1, "Cess");

            const string s4D2 = "4.D.2 Ineligible ITC u/s 16(4) & PoS restricted";
            AddRow(s4D2, "IGST");
            AddRow(s4D2, "CGST");
            AddRow(s4D2, "SGST");
            AddRow(s4D2, "Cess");

            // 5 Non-Taxable Inward Supplies
            const string s5Gst = "5.1 From supplier under composition, Exempt & Nil rated";
            AddRow(s5Gst, "Inter-State");
            AddRow(s5Gst, "Intra-State");

            const string s5NonGst = "5.2 Non-GST inward supply";
            AddRow(s5NonGst, "Inter-State");
            AddRow(s5NonGst, "Intra-State");

            // 6.1 Payment of Tax
            const string s61Oth = "6.1 Total Liability (Other than reverse charge)";
            AddRow(s61Oth, "IGST");
            AddRow(s61Oth, "CGST");
            AddRow(s61Oth, "SGST");
            AddRow(s61Oth, "Cess");

            const string s61Rev = "6.1 Total Liability (Reverse Charge)";
            AddRow(s61Rev, "IGST");
            AddRow(s61Rev, "CGST");
            AddRow(s61Rev, "SGST");
            AddRow(s61Rev, "Cess");

            const string s61TotLiab = "6.1 Total Tax Liability";
            AddRow(s61TotLiab, "IGST", isSummary: true);
            AddRow(s61TotLiab, "CGST", isSummary: true);
            AddRow(s61TotLiab, "SGST", isSummary: true);
            AddRow(s61TotLiab, "Cess", isSummary: true);

            const string s61Itc = "6.1 Total Paid through ITC";
            AddRow(s61Itc, "IGST", isSummary: true);
            AddRow(s61Itc, "CGST", isSummary: true);
            AddRow(s61Itc, "SGST", isSummary: true);
            AddRow(s61Itc, "Cess", isSummary: true);

            const string s61Cash = "6.1 Tax paid in cash";
            AddRow(s61Cash, "IGST");
            AddRow(s61Cash, "CGST");
            AddRow(s61Cash, "SGST");
            AddRow(s61Cash, "Cess");

            const string s61Intr = "6.1 Interest paid in cash";
            AddRow(s61Intr, "IGST");
            AddRow(s61Intr, "CGST");
            AddRow(s61Intr, "SGST");
            AddRow(s61Intr, "Cess");

            const string s61Fee = "6.1 Late Fee paid in cash";
            AddRow(s61Fee, "CGST");
            AddRow(s61Fee, "SGST");

            void AddVal(string section, string typeName, int m, decimal val)
            {
                if (m >= 0 && m < 12 && map.TryGetValue($"{section}|{typeName}", out var row))
                {
                    row.MonthlyValues[m] += val;
                    row.HasMonthData[m] = true;
                }
            }

            foreach (var file in files)
            {
                int m = Get12MonthSlot(file, startYear);
                if (m < 0 || m >= 12) continue;

                // Mark all rows for this GSTIN as having data for month m so 0.00 / "-" is displayed for filed months and blank for unfiled months
                foreach (var r in rows)
                {
                    r.HasMonthData[m] = true;
                }

                try
                {
                    using var stream = AnnualReportScannerService.OpenFileStream(file);
                    using var doc = JsonDocument.Parse(stream);
                    var root = UnwrapData(doc.RootElement);

                    if (root.TryGetProperty("sup_details", out var sup) && sup.ValueKind == JsonValueKind.Object)
                    {
                        if (sup.TryGetProperty("osup_det", out var osupDet))
                        {
                            AddVal(s31A, "Taxable", m, GetDec(osupDet, "txval"));
                            AddVal(s31A, "IGST", m, GetDec(osupDet, "iamt"));
                            AddVal(s31A, "CGST", m, GetDec(osupDet, "camt"));
                            AddVal(s31A, "SGST", m, GetDec(osupDet, "samt"));
                            AddVal(s31A, "Cess", m, GetDec(osupDet, "csamt"));
                        }
                        if (sup.TryGetProperty("osup_zero", out var osupZero))
                        {
                            AddVal(s31B, "Taxable", m, GetDec(osupZero, "txval"));
                            AddVal(s31B, "IGST", m, GetDec(osupZero, "iamt"));
                            AddVal(s31B, "Cess", m, GetDec(osupZero, "csamt"));
                        }
                        if (sup.TryGetProperty("osup_nil_exmp", out var osupNil))
                        {
                            AddVal(s31C, "Taxable", m, GetDec(osupNil, "txval"));
                        }
                        if (sup.TryGetProperty("isup_rev", out var isupRev))
                        {
                            AddVal(s31D, "Taxable", m, GetDec(isupRev, "txval"));
                            AddVal(s31D, "IGST", m, GetDec(isupRev, "iamt"));
                            AddVal(s31D, "CGST", m, GetDec(isupRev, "camt"));
                            AddVal(s31D, "SGST", m, GetDec(isupRev, "samt"));
                            AddVal(s31D, "Cess", m, GetDec(isupRev, "csamt"));
                        }
                        if (sup.TryGetProperty("osup_nongst", out var osupNonGst))
                        {
                            AddVal(s31E, "Taxable", m, GetDec(osupNonGst, "txval"));
                        }
                    }

                    if (root.TryGetProperty("eco_dtls", out var eco) && eco.ValueKind == JsonValueKind.Object)
                    {
                        if (eco.TryGetProperty("eco_sup", out var ecoSup))
                        {
                            AddVal(s311_i, "Taxable", m, GetDec(ecoSup, "txval"));
                            AddVal(s311_i, "IGST", m, GetDec(ecoSup, "iamt"));
                            AddVal(s311_i, "CGST", m, GetDec(ecoSup, "camt"));
                            AddVal(s311_i, "SGST", m, GetDec(ecoSup, "samt"));
                            AddVal(s311_i, "Cess", m, GetDec(ecoSup, "csamt"));
                        }
                        if (eco.TryGetProperty("eco_reg_sup", out var ecoReg))
                        {
                            AddVal(s311_ii, "Taxable", m, GetDec(ecoReg, "txval"));
                        }
                    }

                    if (root.TryGetProperty("itc_elg", out var itcElg) && itcElg.ValueKind == JsonValueKind.Object)
                    {
                        decimal avlI = 0, avlC = 0, avlS = 0, avlCs = 0;
                        decimal revI = 0, revC = 0, revS = 0, revCs = 0;

                        if (itcElg.TryGetProperty("itc_avl", out var itcAvl) && itcAvl.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in itcAvl.EnumerateArray())
                            {
                                string ty = GetStr(item, "ty").ToUpperInvariant();
                                decimal iamt = GetDec(item, "iamt");
                                decimal camt = GetDec(item, "camt");
                                decimal samt = GetDec(item, "samt");
                                decimal csamt = GetDec(item, "csamt");

                                avlI += iamt; avlC += camt; avlS += samt; avlCs += csamt;

                                switch (ty)
                                {
                                    case "IMPG":
                                        AddVal(s4A1, "IGST", m, iamt);
                                        AddVal(s4A1, "Cess", m, csamt);
                                        break;
                                    case "IMPS":
                                        AddVal(s4A2, "IGST", m, iamt);
                                        AddVal(s4A2, "Cess", m, csamt);
                                        break;
                                    case "ISRC":
                                        AddVal(s4A3, "IGST", m, iamt);
                                        AddVal(s4A3, "CGST", m, camt);
                                        AddVal(s4A3, "SGST", m, samt);
                                        AddVal(s4A3, "Cess", m, csamt);
                                        break;
                                    case "ISD":
                                        AddVal(s4A4, "IGST", m, iamt);
                                        AddVal(s4A4, "CGST", m, camt);
                                        AddVal(s4A4, "SGST", m, samt);
                                        AddVal(s4A4, "Cess", m, csamt);
                                        break;
                                    case "OTH":
                                        AddVal(s4A5, "IGST", m, iamt);
                                        AddVal(s4A5, "CGST", m, camt);
                                        AddVal(s4A5, "SGST", m, samt);
                                        AddVal(s4A5, "Cess", m, csamt);
                                        break;
                                }
                            }
                        }

                        if (itcElg.TryGetProperty("itc_rev", out var itcRev) && itcRev.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in itcRev.EnumerateArray())
                            {
                                string ty = GetStr(item, "ty").ToUpperInvariant();
                                string sec = ty == "RUL" ? s4B1 : s4B2;
                                decimal iamt = GetDec(item, "iamt");
                                decimal camt = GetDec(item, "camt");
                                decimal samt = GetDec(item, "samt");
                                decimal csamt = GetDec(item, "csamt");

                                revI += iamt; revC += camt; revS += samt; revCs += csamt;

                                AddVal(sec, "IGST", m, iamt);
                                AddVal(sec, "CGST", m, camt);
                                AddVal(sec, "SGST", m, samt);
                                AddVal(sec, "Cess", m, csamt);
                            }
                        }

                        if (itcElg.TryGetProperty("itc_net", out var itcNet) && itcNet.ValueKind == JsonValueKind.Object)
                        {
                            decimal netI = GetDec(itcNet, "iamt");
                            decimal netC = GetDec(itcNet, "camt");
                            decimal netS = GetDec(itcNet, "samt");
                            decimal netCs = GetDec(itcNet, "csamt");

                            if (netI == 0 && netC == 0 && netS == 0 && netCs == 0 && (avlI != 0 || avlC != 0 || avlS != 0 || avlCs != 0 || revI != 0 || revC != 0 || revS != 0 || revCs != 0))
                            {
                                netI = avlI - revI;
                                netC = avlC - revC;
                                netS = avlS - revS;
                                netCs = avlCs - revCs;
                            }

                            AddVal(sNetItc, "IGST", m, netI);
                            AddVal(sNetItc, "CGST", m, netC);
                            AddVal(sNetItc, "SGST", m, netS);
                            AddVal(sNetItc, "Cess", m, netCs);
                        }
                        else
                        {
                            AddVal(sNetItc, "IGST", m, avlI - revI);
                            AddVal(sNetItc, "CGST", m, avlC - revC);
                            AddVal(sNetItc, "SGST", m, avlS - revS);
                            AddVal(sNetItc, "Cess", m, avlCs - revCs);
                        }

                        if (itcElg.TryGetProperty("itc_inelg", out var itcInelg) && itcInelg.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in itcInelg.EnumerateArray())
                            {
                                string ty = GetStr(item, "ty").ToUpperInvariant();
                                string sec = ty == "RUL" ? s4D1 : s4D2;
                                AddVal(sec, "IGST", m, GetDec(item, "iamt"));
                                AddVal(sec, "CGST", m, GetDec(item, "camt"));
                                AddVal(sec, "SGST", m, GetDec(item, "samt"));
                                AddVal(sec, "Cess", m, GetDec(item, "csamt"));
                            }
                        }
                    }

                    if (root.TryGetProperty("inward_sup", out var inwSup) &&
                        inwSup.ValueKind == JsonValueKind.Object &&
                        inwSup.TryGetProperty("isup_details", out var isupDetails) &&
                        isupDetails.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var item in isupDetails.EnumerateArray())
                        {
                            string ty = GetStr(item, "ty").ToUpperInvariant();
                            if (ty == "GST")
                            {
                                AddVal(s5Gst, "Inter-State", m, GetDec(item, "inter"));
                                AddVal(s5Gst, "Intra-State", m, GetDec(item, "intra"));
                            }
                            else if (ty == "NONGST")
                            {
                                AddVal(s5NonGst, "Inter-State", m, GetDec(item, "inter"));
                                AddVal(s5NonGst, "Intra-State", m, GetDec(item, "intra"));
                            }
                        }
                    }

                    Extract3BPaymentOfTax(root, m, AddVal, s61Oth, s61Rev, s61TotLiab, s61Itc, s61Cash, s61Intr, s61Fee);
                }
                catch
                {
                    // Continue
                }
            }

            return rows;
        }

        private static void Extract3BPaymentOfTax(
            JsonElement root,
            int m,
            Action<string, string, int, decimal> addVal,
            string s61Oth,
            string s61Rev,
            string s61TotLiab,
            string s61Itc,
            string s61Cash,
            string s61Intr,
            string s61Fee)
        {
            decimal othIgst = 0, othCgst = 0, othSgst = 0, othCess = 0;
            decimal revIgst = 0, revCgst = 0, revSgst = 0, revCess = 0;
            decimal itcI = 0, itcC = 0, itcS = 0, itcCs = 0;
            decimal cshIgst = 0, cshCgst = 0, cshSgst = 0, cshCess = 0;
            decimal intrCshI = 0, intrCshC = 0, intrCshS = 0, intrCshCs = 0;
            decimal feeCshC = 0, feeCshS = 0;

            // Schema 1: GST Portal taxpayble.returnsDbCdredList
            if (root.TryGetProperty("taxpayble", out var taxPayble) &&
                taxPayble.ValueKind == JsonValueKind.Object &&
                taxPayble.TryGetProperty("returnsDbCdredList", out var dbList) &&
                dbList.ValueKind == JsonValueKind.Object)
            {
                if (dbList.TryGetProperty("tax_pay", out var taxPayArr) && taxPayArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var entry in taxPayArr.EnumerateArray())
                    {
                        int trancd = entry.TryGetProperty("trancd", out var tc) && tc.ValueKind == JsonValueKind.Number ? tc.GetInt32() : 0;
                        decimal iTx = GetNestedDec(entry, "igst", "tx");
                        decimal cTx = GetNestedDec(entry, "cgst", "tx");
                        decimal sTx = GetNestedDec(entry, "sgst", "tx");
                        decimal csTx = GetNestedDec(entry, "cess", "tx");

                        if (trancd == 30003)
                        {
                            revIgst += iTx; revCgst += cTx; revSgst += sTx; revCess += csTx;
                        }
                        else
                        {
                            othIgst += iTx; othCgst += cTx; othSgst += sTx; othCess += csTx;
                        }
                    }
                }

                if (dbList.TryGetProperty("tax_paid", out var taxPaid) && taxPaid.ValueKind == JsonValueKind.Object)
                {
                    if (taxPaid.TryGetProperty("pd_by_itc", out var pdItc))
                    {
                        if (pdItc.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var entry in pdItc.EnumerateArray())
                            {
                                AccumulateItcPaidEntry(entry, ref itcI, ref itcC, ref itcS, ref itcCs);
                            }
                        }
                        else if (pdItc.ValueKind == JsonValueKind.Object)
                        {
                            AccumulateItcPaidEntry(pdItc, ref itcI, ref itcC, ref itcS, ref itcCs);
                        }
                    }

                    if (taxPaid.TryGetProperty("pd_by_cash", out var pdCash) && pdCash.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var entry in pdCash.EnumerateArray())
                        {
                            cshIgst += GetNestedDec(entry, "igst", "tx") + GetDec(entry, "i_pd");
                            cshCgst += GetNestedDec(entry, "cgst", "tx") + GetDec(entry, "c_pd");
                            cshSgst += GetNestedDec(entry, "sgst", "tx") + GetDec(entry, "s_pd");
                            cshCess += GetNestedDec(entry, "cess", "tx") + GetDec(entry, "cs_pd");

                            intrCshI += GetNestedDec(entry, "igst", "intr") + GetDec(entry, "i_intrpd");
                            intrCshC += GetNestedDec(entry, "cgst", "intr") + GetDec(entry, "c_intrpd");
                            intrCshS += GetNestedDec(entry, "sgst", "intr") + GetDec(entry, "s_intrpd");
                            intrCshCs += GetNestedDec(entry, "cess", "intr") + GetDec(entry, "cs_intrpd");

                            feeCshC += GetNestedDec(entry, "cgst", "fee") + GetDec(entry, "c_lfeepd");
                            feeCshS += GetNestedDec(entry, "sgst", "fee") + GetDec(entry, "s_lfeepd");
                        }
                    }
                }
            }
            // Schema 2: GSTN API tx_pmt
            else if (root.TryGetProperty("tx_pmt", out var txPmt) && txPmt.ValueKind == JsonValueKind.Object)
            {
                if (txPmt.TryGetProperty("tx_py", out var txPyArr) && txPyArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var entry in txPyArr.EnumerateArray())
                    {
                        int transTyp = entry.TryGetProperty("trans_typ", out var tt) && tt.ValueKind == JsonValueKind.Number ? tt.GetInt32() : 0;
                        string desc = GetStr(entry, "tran_desc");
                        bool isRcm = transTyp == 30003 || desc.Contains("Reverse", StringComparison.OrdinalIgnoreCase);

                        decimal iTx = GetNestedDec(entry, "igst", "tx") + GetDec(entry, "i_tx");
                        decimal cTx = GetNestedDec(entry, "cgst", "tx") + GetDec(entry, "c_tx");
                        decimal sTx = GetNestedDec(entry, "sgst", "tx") + GetDec(entry, "s_tx");
                        decimal csTx = GetNestedDec(entry, "cess", "tx") + GetDec(entry, "cs_tx");

                        if (isRcm)
                        {
                            revIgst += iTx; revCgst += cTx; revSgst += sTx; revCess += csTx;
                        }
                        else
                        {
                            othIgst += iTx; othCgst += cTx; othSgst += sTx; othCess += csTx;
                        }
                    }
                }

                if (txPmt.TryGetProperty("pditc", out var pdItc2))
                {
                    if (pdItc2.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var entry in pdItc2.EnumerateArray())
                        {
                            AccumulateItcPaidEntry(entry, ref itcI, ref itcC, ref itcS, ref itcCs);
                        }
                    }
                    else if (pdItc2.ValueKind == JsonValueKind.Object)
                    {
                        AccumulateItcPaidEntry(pdItc2, ref itcI, ref itcC, ref itcS, ref itcCs);
                    }
                }

                if (txPmt.TryGetProperty("pdcash", out var pdCash2) && pdCash2.ValueKind == JsonValueKind.Array)
                {
                    foreach (var entry in pdCash2.EnumerateArray())
                    {
                        cshIgst += GetDec(entry, "i_pd") + GetNestedDec(entry, "igst", "tx");
                        cshCgst += GetDec(entry, "c_pd") + GetNestedDec(entry, "cgst", "tx");
                        cshSgst += GetDec(entry, "s_pd") + GetNestedDec(entry, "sgst", "tx");
                        cshCess += GetDec(entry, "cs_pd") + GetNestedDec(entry, "cess", "tx");

                        intrCshI += GetDec(entry, "i_intrpd") + GetNestedDec(entry, "igst", "intr");
                        intrCshC += GetDec(entry, "c_intrpd") + GetNestedDec(entry, "cgst", "intr");
                        intrCshS += GetDec(entry, "s_intrpd") + GetNestedDec(entry, "sgst", "intr");
                        intrCshCs += GetDec(entry, "cs_intrpd") + GetNestedDec(entry, "cess", "intr");

                        feeCshC += GetDec(entry, "c_lfeepd") + GetNestedDec(entry, "cgst", "fee");
                        feeCshS += GetDec(entry, "s_lfeepd") + GetNestedDec(entry, "sgst", "fee");
                    }
                }
            }

            addVal(s61Oth, "IGST", m, othIgst);
            addVal(s61Oth, "CGST", m, othCgst);
            addVal(s61Oth, "SGST", m, othSgst);
            addVal(s61Oth, "Cess", m, othCess);

            addVal(s61Rev, "IGST", m, revIgst);
            addVal(s61Rev, "CGST", m, revCgst);
            addVal(s61Rev, "SGST", m, revSgst);
            addVal(s61Rev, "Cess", m, revCess);

            addVal(s61TotLiab, "IGST", m, othIgst + revIgst);
            addVal(s61TotLiab, "CGST", m, othCgst + revCgst);
            addVal(s61TotLiab, "SGST", m, othSgst + revSgst);
            addVal(s61TotLiab, "Cess", m, othCess + revCess);

            addVal(s61Itc, "IGST", m, itcI);
            addVal(s61Itc, "CGST", m, itcC);
            addVal(s61Itc, "SGST", m, itcS);
            addVal(s61Itc, "Cess", m, itcCs);

            addVal(s61Cash, "IGST", m, cshIgst);
            addVal(s61Cash, "CGST", m, cshCgst);
            addVal(s61Cash, "SGST", m, cshSgst);
            addVal(s61Cash, "Cess", m, cshCess);

            addVal(s61Intr, "IGST", m, intrCshI);
            addVal(s61Intr, "CGST", m, intrCshC);
            addVal(s61Intr, "SGST", m, intrCshS);
            addVal(s61Intr, "Cess", m, intrCshCs);

            addVal(s61Fee, "CGST", m, feeCshC);
            addVal(s61Fee, "SGST", m, feeCshS);
        }

        private static void AccumulateItcPaidEntry(
            JsonElement entry,
            ref decimal itcI,
            ref decimal itcC,
            ref decimal itcS,
            ref decimal itcCs)
        {
            itcI += GetDec(entry, "igst_igst_amt") + GetDec(entry, "igst_cgst_amt") + GetDec(entry, "igst_sgst_amt")
                  + GetDec(entry, "i_pdi") + GetDec(entry, "i_pdc") + GetDec(entry, "i_pds");
            itcC += GetDec(entry, "cgst_igst_amt") + GetDec(entry, "cgst_cgst_amt")
                  + GetDec(entry, "c_pdi") + GetDec(entry, "c_pdc");
            itcS += GetDec(entry, "sgst_igst_amt") + GetDec(entry, "sgst_sgst_amt")
                  + GetDec(entry, "s_pdi") + GetDec(entry, "s_pds");
            itcCs += GetDec(entry, "cess_cess_amt") + GetDec(entry, "cs_pdcs");
        }

        #endregion

        #region 2. GSTR-1 Builder (24 Monthly Columns: Apr YYYY .. Mar YYYY+2 + Total, with Type = Taxable / IGST / CGST / SGST / Cess)

        private List<AnnualSheetRow> BuildGstr1RowsForGstin(
            string gstin,
            string gstinDisplay,
            List<AnnualScannedFile> files,
            int startYear,
            DateTime fyStartDate)
        {
            var map = new Dictionary<string, AnnualSheetRow>(StringComparer.OrdinalIgnoreCase);

            void AddSectionTaxRows(string section, int baseOrder, int m, decimal txval, decimal iamt, decimal camt, decimal samt, decimal csamt)
            {
                AddMetric(map, 24, gstin, gstinDisplay, section, "Taxable", baseOrder, 1, m, txval);
                AddMetric(map, 24, gstin, gstinDisplay, section, "IGST", baseOrder, 2, m, iamt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "CGST", baseOrder, 3, m, camt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "SGST", baseOrder, 4, m, samt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "Cess", baseOrder, 5, m, csamt);
            }

            // Ensure baseline B2B section exists
            AddSectionTaxRows("B2B Regular Supplies", 10, -1, 0, 0, 0, 0, 0);
            AddSectionTaxRows("B2CS Small Supplies", 30, -1, 0, 0, 0, 0, 0);
            AddSectionTaxRows("CDNR Credit/Debit Notes", 50, -1, 0, 0, 0, 0, 0);
            AddSectionTaxRows("NIL / Exempt Supplies", 70, -1, 0, 0, 0, 0, 0);

            var activeSlots = new HashSet<int>();

            foreach (var file in files)
            {
                int m = Get24MonthSlot(file, startYear);
                if (m < 0 || m >= 24) continue;
                activeSlots.Add(m);

                try
                {
                    using var stream = AnnualReportScannerService.OpenFileStream(file);
                    using var doc = JsonDocument.Parse(stream);
                    var root = UnwrapData(doc.RootElement);

                    AccumulateR1B2B(root, "b2b", "B2B", 10, m, fyStartDate, AddSectionTaxRows);
                    AccumulateR1B2B(root, "b2ba", "B2B-A", 15, m, fyStartDate, AddSectionTaxRows);
                    AccumulateR1InvoiceSection(root, "b2cl", "B2CL Large Supplies", 20, m, AddSectionTaxRows);
                    AccumulateR1InvoiceSection(root, "b2cla", "B2CL-A Amendments", 25, m, AddSectionTaxRows);
                    AccumulateR1FlatSection(root, "b2cs", "B2CS Small Supplies", 30, m, 1m, AddSectionTaxRows);
                    AccumulateR1FlatSection(root, "b2csa", "B2CS-A Amendments", 35, m, 1m, AddSectionTaxRows);
                    AccumulateR1InvoiceSection(root, "exp", "EXP Exports", 40, m, AddSectionTaxRows);
                    AccumulateR1InvoiceSection(root, "expa", "EXP-A Amendments", 45, m, AddSectionTaxRows);
                    AccumulateR1NotesSection(root, "cdnr", "CDNR Credit/Debit Notes", 50, m, fyStartDate, AddSectionTaxRows);
                    AccumulateR1NotesSection(root, "cdnra", "CDNR-A Note Amendments", 55, m, fyStartDate, AddSectionTaxRows);
                    AccumulateR1NotesSection(root, "cdnur", "CDNUR Unregistered Notes", 60, m, fyStartDate, AddSectionTaxRows);
                    AccumulateR1NotesSection(root, "cdnura", "CDNUR-A Amendments", 65, m, fyStartDate, AddSectionTaxRows);

                    if (root.TryGetProperty("nil", out var nilObj) && nilObj.ValueKind == JsonValueKind.Object &&
                        nilObj.TryGetProperty("inv", out var nilInv) && nilInv.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var entry in nilInv.EnumerateArray())
                        {
                            decimal nilAmt = GetDec(entry, "nil_amt") + GetDec(entry, "expt_amt") + GetDec(entry, "ngsup_amt");
                            AddSectionTaxRows("NIL / Exempt Supplies", 70, m, nilAmt, 0, 0, 0, 0);
                        }
                    }

                    AccumulateR1FlatSection(root, "at", "AT / TXPD Advances", 80, m, 1m, AddSectionTaxRows);
                    AccumulateR1FlatSection(root, "ata", "AT-A / TXPD-A Amendments", 82, m, 1m, AddSectionTaxRows);
                    AccumulateR1FlatSection(root, "txpd", "AT / TXPD Advances", 80, m, -1m, AddSectionTaxRows);
                    AccumulateR1FlatSection(root, "txpda", "AT-A / TXPD-A Amendments", 82, m, -1m, AddSectionTaxRows);

                    if (root.TryGetProperty("supeco", out var supEco) && supEco.ValueKind == JsonValueKind.Object)
                    {
                        AccumulateR1FlatSection(supEco, "clttx", "ECO Supplies (u/s 52)", 85, m, 1m, AddSectionTaxRows);
                        AccumulateR1FlatSection(supEco, "paytx", "ECO Supplies (u/s 9(5))", 86, m, 1m, AddSectionTaxRows);
                    }
                }
                catch
                {
                    // Continue
                }
            }

            var ordered = map.Values
                .OrderBy(r => r.SortOrder)
                .ThenBy(r => r.Section)
                .ThenBy(r => r.SubOrder)
                .ToList();

            foreach (var r in ordered)
            {
                foreach (int slot in activeSlots)
                {
                    r.HasMonthData[slot] = true;
                }
            }

            // Add Total Outward Supplies Summary Rows
            string[] heads = { "Taxable", "IGST", "CGST", "SGST", "Cess" };
            for (int i = 0; i < heads.Length; i++)
            {
                string head = heads[i];
                var sumRow = new AnnualSheetRow(24)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = "Total Outward Supplies",
                    Description = head,
                    SortOrder = 900 + i,
                    SubOrder = i + 1,
                    IsSectionSummaryRow = true
                };

                for (int m = 0; m < 24; m++)
                {
                    sumRow.MonthlyValues[m] = ordered
                        .Where(r => r.Description == head)
                        .Sum(r => r.MonthlyValues[m]);
                    sumRow.HasMonthData[m] = activeSlots.Contains(m);
                }
                ordered.Add(sumRow);
            }

            return ordered;
        }

        private static void AccumulateR1B2B(
            JsonElement root,
            string propName,
            string baseSection,
            int baseSortOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var party in arr.EnumerateArray())
            {
                if (!party.TryGetProperty("inv", out var invArr) || invArr.ValueKind != JsonValueKind.Array) continue;
                foreach (var inv in invArr.EnumerateArray())
                {
                    string rchrg = GetStr(inv, "rchrg").ToUpperInvariant();
                    string invTyp = GetStr(inv, "inv_typ").ToUpperInvariant();
                    string idt = GetStr(inv, "idt");
                    bool isPrevFy = IsDateBefore(idt, fyStartDate);
                    string prevSuffix = isPrevFy ? " (prev-FY)" : string.Empty;

                    string section;
                    int order;
                    if (rchrg == "Y")
                    {
                        section = $"{baseSection} Reverse Charge Supplies{prevSuffix}";
                        order = baseSortOrder + 1;
                    }
                    else if (invTyp == "SEWP" || invTyp == "SEWOP" || invTyp == "DE")
                    {
                        section = $"{baseSection} SEZ / Deemed Exports{prevSuffix}";
                        order = baseSortOrder + 2;
                    }
                    else
                    {
                        section = $"{baseSection} Regular Supplies{prevSuffix}";
                        order = baseSortOrder;
                    }

                    ExtractInvoiceItems(inv, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                    addSectionTaxRows(section, order, m, txval, iamt, camt, samt, csamt);
                }
            }
        }

        private static void AccumulateR1InvoiceSection(
            JsonElement root,
            string propName,
            string section,
            int sortOrder,
            int m,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var group in arr.EnumerateArray())
            {
                if (group.TryGetProperty("inv", out var invArr) && invArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var inv in invArr.EnumerateArray())
                    {
                        ExtractInvoiceItems(inv, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                        addSectionTaxRows(section, sortOrder, m, txval, iamt, camt, samt, csamt);
                    }
                }
                else
                {
                    ExtractInvoiceItems(group, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                    addSectionTaxRows(section, sortOrder, m, txval, iamt, camt, samt, csamt);
                }
            }
        }

        private static void AccumulateR1FlatSection(
            JsonElement root,
            string propName,
            string section,
            int sortOrder,
            int m,
            decimal sign,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var entry in arr.EnumerateArray())
            {
                ExtractInvoiceItems(entry, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                addSectionTaxRows(section, sortOrder, m, sign * txval, sign * iamt, sign * camt, sign * samt, sign * csamt);
            }
        }

        private static void AccumulateR1NotesSection(
            JsonElement root,
            string propName,
            string baseSection,
            int sortOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;

            foreach (var party in arr.EnumerateArray())
            {
                if (party.TryGetProperty("nt", out var ntArr) && ntArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var nt in ntArr.EnumerateArray())
                    {
                        ProcessSingleNoteR1(nt, baseSection, sortOrder, m, fyStartDate, addSectionTaxRows);
                    }
                }
                else
                {
                    ProcessSingleNoteR1(party, baseSection, sortOrder, m, fyStartDate, addSectionTaxRows);
                }
            }
        }

        private static void ProcessSingleNoteR1(
            JsonElement nt,
            string baseSection,
            int sortOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            string ntty = GetStr(nt, "ntty").ToUpperInvariant();
            decimal sign = (ntty == "C" || ntty == "CR") ? -1m : 1m;
            string idt = GetStr(nt, "idt");
            if (string.IsNullOrWhiteSpace(idt)) idt = GetStr(nt, "nt_dt");
            bool isPrevFy = IsDateBefore(idt, fyStartDate);
            string section = isPrevFy ? $"{baseSection} (prev-FY)" : baseSection;

            ExtractInvoiceItems(nt, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
            addSectionTaxRows(section, sortOrder + (isPrevFy ? 1 : 0), m, sign * txval, sign * iamt, sign * camt, sign * samt, sign * csamt);
        }

        #endregion

        #region 3. GSTR-2A and GSTR-2B Builders (24 Monthly Columns: Apr YYYY .. Mar YYYY+2 + Total)

        private List<AnnualSheetRow> BuildGstr2ARowsForGstin(
            string gstin,
            string gstinDisplay,
            List<AnnualScannedFile> files,
            int startYear,
            DateTime fyStartDate)
        {
            var map = new Dictionary<string, AnnualSheetRow>(StringComparer.OrdinalIgnoreCase);
            var activeSlots = new HashSet<int>();

            void AddSectionTaxRows(string section, int baseOrder, int m, decimal txval, decimal iamt, decimal camt, decimal samt, decimal csamt)
            {
                AddMetric(map, 24, gstin, gstinDisplay, section, "Taxable", baseOrder, 1, m, txval);
                AddMetric(map, 24, gstin, gstinDisplay, section, "IGST", baseOrder, 2, m, iamt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "CGST", baseOrder, 3, m, camt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "SGST", baseOrder, 4, m, samt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "Cess", baseOrder, 5, m, csamt);
            }

            AddSectionTaxRows("B2B", 10, -1, 0, 0, 0, 0, 0);

            foreach (var file in files)
            {
                int m = Get24MonthSlot(file, startYear);
                if (m < 0 || m >= 24) continue;
                activeSlots.Add(m);

                try
                {
                    using var stream = AnnualReportScannerService.OpenFileStream(file);
                    using var doc = JsonDocument.Parse(stream);
                    var root = UnwrapData(doc.RootElement);

                    Accumulate2APartyInvoices(root, "b2b", "B2B", 10, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2APartyInvoices(root, "b2ba", "B2B-A", 20, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2APartyNotes(root, "cdn", "CDNR", 30, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2APartyNotes(root, "cdna", "CDNR-A", 40, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2APartyInvoices(root, "isd", "ISD", 50, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2APartyInvoices(root, "isda", "ISD-A", 55, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2AImpg(root, "impg", "IMPG", 60, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2AImpg(root, "impgsez", "IMPGSEZ", 70, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2APartyInvoices(root, "eco", "ECO", 80, m, fyStartDate, AddSectionTaxRows);
                }
                catch
                {
                    // Continue
                }
            }

            return FinalizeInwardSheetRows(gstin, gstinDisplay, 24, map, activeSlots);
        }

        private List<AnnualSheetRow> BuildGstr2BRowsForGstin(
            string gstin,
            string gstinDisplay,
            List<AnnualScannedFile> files,
            int startYear,
            DateTime fyStartDate)
        {
            var map = new Dictionary<string, AnnualSheetRow>(StringComparer.OrdinalIgnoreCase);
            var activeSlots = new HashSet<int>();

            void AddSectionTaxRows(string section, int baseOrder, int m, decimal txval, decimal iamt, decimal camt, decimal samt, decimal csamt)
            {
                AddMetric(map, 24, gstin, gstinDisplay, section, "Taxable", baseOrder, 1, m, txval);
                AddMetric(map, 24, gstin, gstinDisplay, section, "IGST", baseOrder, 2, m, iamt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "CGST", baseOrder, 3, m, camt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "SGST", baseOrder, 4, m, samt);
                AddMetric(map, 24, gstin, gstinDisplay, section, "Cess", baseOrder, 5, m, csamt);
            }

            AddSectionTaxRows("B2B", 10, -1, 0, 0, 0, 0, 0);

            foreach (var file in files)
            {
                int m = Get24MonthSlot(file, startYear);
                if (m < 0 || m >= 24) continue;
                activeSlots.Add(m);

                try
                {
                    using var stream = AnnualReportScannerService.OpenFileStream(file);
                    using var doc = JsonDocument.Parse(stream);
                    var payload = UnwrapData(doc.RootElement);
                    var docData = payload.TryGetProperty("docdata", out var dd) && dd.ValueKind == JsonValueKind.Object ? dd : payload;

                    Accumulate2BPartyInvoices(docData, "b2b", "B2B", 10, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BPartyInvoices(docData, "b2ba", "B2B-A", 20, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BPartyNotes(docData, "cdnr", "CDNR", 30, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BPartyNotes(docData, "cdnra", "CDNR-A", 40, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BPartyInvoices(docData, "isd", "ISD", 50, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BPartyInvoices(docData, "isda", "ISD-A", 55, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BImpg(docData, "impg", "IMPG", 60, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BImpg(docData, "impgsez", "IMPGSEZ", 70, m, fyStartDate, AddSectionTaxRows);
                    Accumulate2BPartyInvoices(docData, "eco", "ECO", 80, m, fyStartDate, AddSectionTaxRows);
                }
                catch
                {
                    // Continue
                }
            }

            return FinalizeInwardSheetRows(gstin, gstinDisplay, 24, map, activeSlots);
        }

        private static List<AnnualSheetRow> FinalizeInwardSheetRows(
            string gstin,
            string gstinDisplay,
            int periodCount,
            Dictionary<string, AnnualSheetRow> map,
            HashSet<int> activeSlots)
        {
            var ordered = map.Values
                .OrderBy(r => r.SortOrder)
                .ThenBy(r => r.Section)
                .ThenBy(r => r.SubOrder)
                .ToList();

            foreach (var r in ordered)
            {
                foreach (int slot in activeSlots)
                {
                    r.HasMonthData[slot] = true;
                }
            }

            string[] heads = { "Taxable", "IGST", "CGST", "SGST", "Cess" };
            for (int i = 0; i < heads.Length; i++)
            {
                string head = heads[i];
                var sumRow = new AnnualSheetRow(periodCount)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = "Total Net ITC (All Eligible Sections)",
                    Description = head,
                    SortOrder = 900 + i,
                    SubOrder = i + 1,
                    IsSectionSummaryRow = true
                };

                for (int m = 0; m < periodCount; m++)
                {
                    sumRow.MonthlyValues[m] = ordered
                        .Where(r => r.Description == head && !r.Section.Contains("(Rejected)", StringComparison.OrdinalIgnoreCase))
                        .Sum(r => r.MonthlyValues[m]);
                    sumRow.HasMonthData[m] = activeSlots.Contains(m);
                }
                ordered.Add(sumRow);
            }

            return ordered;
        }

        private static void Accumulate2APartyInvoices(
            JsonElement root,
            string propName,
            string baseSection,
            int baseOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;
            foreach (var party in arr.EnumerateArray())
            {
                if (!party.TryGetProperty("inv", out var invArr) || invArr.ValueKind != JsonValueKind.Array) continue;
                foreach (var inv in invArr.EnumerateArray())
                {
                    string dt = GetStr(inv, "idt");
                    if (string.IsNullOrWhiteSpace(dt)) dt = GetStr(inv, "dt");
                    bool isPrev = IsDateBefore(dt, fyStartDate);
                    string section = isPrev ? $"{baseSection} (prev-FY)" : baseSection;

                    ExtractInvoiceItems(inv, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                    addSectionTaxRows(section, baseOrder + (isPrev ? 2 : 0), m, txval, iamt, camt, samt, csamt);
                }
            }
        }

        private static void Accumulate2APartyNotes(
            JsonElement root,
            string propName,
            string baseSection,
            int baseOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;
            foreach (var party in arr.EnumerateArray())
            {
                if (!party.TryGetProperty("nt", out var ntArr) || ntArr.ValueKind != JsonValueKind.Array) continue;
                foreach (var nt in ntArr.EnumerateArray())
                {
                    string ntty = GetStr(nt, "ntty").ToUpperInvariant();
                    decimal sign = (ntty == "C" || ntty == "CR") ? -1m : 1m;
                    string dt = GetStr(nt, "idt");
                    if (string.IsNullOrWhiteSpace(dt)) dt = GetStr(nt, "nt_dt");
                    bool isPrev = IsDateBefore(dt, fyStartDate);
                    string section = isPrev ? $"{baseSection} (prev-FY)" : baseSection;

                    ExtractInvoiceItems(nt, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                    addSectionTaxRows(section, baseOrder + (isPrev ? 2 : 0), m, sign * txval, sign * iamt, sign * camt, sign * samt, sign * csamt);
                }
            }
        }

        private static void Accumulate2AImpg(
            JsonElement root,
            string propName,
            string baseSection,
            int baseOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!root.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;
            foreach (var item in arr.EnumerateArray())
            {
                string dt = GetStr(item, "boe_dt");
                bool isPrev = IsDateBefore(dt, fyStartDate);
                string section = isPrev ? $"{baseSection} (prev-FY)" : baseSection;
                ExtractInvoiceItems(item, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                addSectionTaxRows(section, baseOrder + (isPrev ? 2 : 0), m, txval, iamt, camt, samt, csamt);
            }
        }

        private static void Accumulate2BPartyInvoices(
            JsonElement docData,
            string propName,
            string baseSection,
            int baseOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!docData.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;
            foreach (var party in arr.EnumerateArray())
            {
                if (!party.TryGetProperty("inv", out var invArr) || invArr.ValueKind != JsonValueKind.Array) continue;
                foreach (var inv in invArr.EnumerateArray())
                {
                    string dt = GetStr(inv, "dt");
                    if (string.IsNullOrWhiteSpace(dt)) dt = GetStr(inv, "idt");
                    bool isPrev = IsDateBefore(dt, fyStartDate);
                    string itcAvl = GetStr(inv, "itcavl").ToUpperInvariant();
                    bool isRejected = itcAvl == "N";

                    string section = baseSection;
                    int offset = 0;
                    if (isPrev) { section += " (prev-FY)"; offset += 2; }
                    if (isRejected) { section += " (Rejected)"; offset += 4; }

                    ExtractInvoiceItems(inv, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                    addSectionTaxRows(section, baseOrder + offset, m, txval, iamt, camt, samt, csamt);
                }
            }
        }

        private static void Accumulate2BPartyNotes(
            JsonElement docData,
            string propName,
            string baseSection,
            int baseOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!docData.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;
            foreach (var party in arr.EnumerateArray())
            {
                if (!party.TryGetProperty("nt", out var ntArr) || ntArr.ValueKind != JsonValueKind.Array) continue;
                foreach (var nt in ntArr.EnumerateArray())
                {
                    string suptyp = GetStr(nt, "ntty").ToUpperInvariant();
                    if (string.IsNullOrWhiteSpace(suptyp)) suptyp = GetStr(nt, "typ").ToUpperInvariant();
                    decimal sign = (suptyp == "C" || suptyp == "CR") ? -1m : 1m;

                    string dt = GetStr(nt, "dt");
                    if (string.IsNullOrWhiteSpace(dt)) dt = GetStr(nt, "idt");
                    bool isPrev = IsDateBefore(dt, fyStartDate);
                    string itcAvl = GetStr(nt, "itcavl").ToUpperInvariant();
                    bool isRejected = itcAvl == "N";

                    string section = baseSection;
                    int offset = 0;
                    if (isPrev) { section += " (prev-FY)"; offset += 2; }
                    if (isRejected) { section += " (Rejected)"; offset += 4; }

                    ExtractInvoiceItems(nt, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                    addSectionTaxRows(section, baseOrder + offset, m, sign * txval, sign * iamt, sign * camt, sign * samt, sign * csamt);
                }
            }
        }

        private static void Accumulate2BImpg(
            JsonElement docData,
            string propName,
            string baseSection,
            int baseOrder,
            int m,
            DateTime fyStartDate,
            Action<string, int, int, decimal, decimal, decimal, decimal, decimal> addSectionTaxRows)
        {
            if (!docData.TryGetProperty(propName, out var arr) || arr.ValueKind != JsonValueKind.Array) return;
            foreach (var item in arr.EnumerateArray())
            {
                string dt = GetStr(item, "boedt");
                if (string.IsNullOrWhiteSpace(dt)) dt = GetStr(item, "boe_dt");
                bool isPrev = IsDateBefore(dt, fyStartDate);
                string section = isPrev ? $"{baseSection} (prev-FY)" : baseSection;
                ExtractInvoiceItems(item, out decimal txval, out decimal iamt, out decimal camt, out decimal samt, out decimal csamt);
                addSectionTaxRows(section, baseOrder + (isPrev ? 2 : 0), m, txval, iamt, camt, samt, csamt);
            }
        }

        private static void AddMetric(
            Dictionary<string, AnnualSheetRow> map,
            int periodCount,
            string gstin,
            string gstinDisplay,
            string section,
            string typeName,
            int sortOrder,
            int subOrder,
            int m,
            decimal val)
        {
            string key = $"{section}|{typeName}";
            if (!map.TryGetValue(key, out var row))
            {
                row = new AnnualSheetRow(periodCount)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = section,
                    Description = typeName,
                    SortOrder = sortOrder,
                    SubOrder = subOrder
                };
                map[key] = row;
            }
            if (m >= 0 && m < periodCount)
            {
                row.MonthlyValues[m] += val;
                row.HasMonthData[m] = true;
            }
        }

        #endregion

        #region 4. Reconciliation Builders (12 Monthly Columns: Apr YYYY .. Mar YYYY+1 + Total)

        private static List<AnnualSheetRow> Build3BVs1ForGstin(
            string gstin,
            string gstinDisplay,
            List<AnnualSheetRow> rows3B,
            List<AnnualSheetRow> rows1,
            List<AnnualReconciliationAlert> alerts)
        {
            var result = new List<AnnualSheetRow>();
            string[] heads = { "Taxable", "IGST", "CGST", "SGST", "Cess" };
            int order = 1;

            foreach (var head in heads)
            {
                var r3B = new AnnualSheetRow(12)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = $"Outward Supplies as per GSTR-3B ({head})",
                    Description = head,
                    SortOrder = order++,
                    IsSectionSummaryRow = true
                };

                var r1 = new AnnualSheetRow(12)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = $"Outward Supplies as per GSTR-1 ({head})",
                    Description = head,
                    SortOrder = order++,
                    IsSectionSummaryRow = true
                };

                var rDiff = new AnnualSheetRow(12)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = $"Difference (GSTR-3B minus GSTR-1)",
                    Description = head,
                    SortOrder = order++,
                    IsDifferenceRow = true
                };

                var r1Summary = rows1.FirstOrDefault(r =>
                    r.Section == "Total Outward Supplies" && r.Description == head);

                for (int m = 0; m < 12; m++)
                {
                    decimal val3B = head == "Taxable"
                        ? Get3BVal(rows3B, "3.1.A Outward taxable supplies (excluding zero rated)", "Taxable", m, out bool h1)
                          + Get3BVal(rows3B, "3.1.B Outward taxable supplies (zero rated)", "Taxable", m, out _)
                          + Get3BVal(rows3B, "3.1.C Other outward supplies (nil-rated, exempted)", "Taxable", m, out _)
                          + Get3BVal(rows3B, "3.1.E Non-GST outward supplies", "Taxable", m, out _)
                        : Get3BVal(rows3B, "3.1.A Outward taxable supplies (excluding zero rated)", head, m, out h1)
                          + Get3BVal(rows3B, "3.1.B Outward taxable supplies (zero rated)", head, m, out _);

                    bool has1 = r1Summary != null && m < r1Summary.HasMonthData.Length && r1Summary.HasMonthData[m];
                    decimal val1 = r1Summary != null && m < r1Summary.MonthlyValues.Length ? r1Summary.MonthlyValues[m] : 0m;

                    bool anyData = h1 || has1;
                    r3B.MonthlyValues[m] = val3B;
                    r3B.HasMonthData[m] = anyData;

                    r1.MonthlyValues[m] = val1;
                    r1.HasMonthData[m] = anyData;

                    rDiff.MonthlyValues[m] = val3B - val1;
                    rDiff.HasMonthData[m] = anyData;
                }

                result.Add(r3B);
                result.Add(r1);
                result.Add(rDiff);

                if (Math.Abs(rDiff.TotalValue) >= 1m)
                {
                    alerts.Add(new AnnualReconciliationAlert
                    {
                        GstinDisplay = gstinDisplay,
                        ComparisonSheet = "GSTR-3B vs GSTR-1",
                        MetricName = $"Outward {head} Difference",
                        DifferenceAmount = rDiff.TotalValue
                    });
                }
            }

            return result;
        }

        private static List<AnnualSheetRow> Build3BVs2AOr2BForGstin(
            string gstin,
            string gstinDisplay,
            string targetReturnName,
            List<AnnualSheetRow> rows3B,
            List<AnnualSheetRow> rowsInward,
            List<AnnualReconciliationAlert> alerts)
        {
            var result = new List<AnnualSheetRow>();
            string[] taxHeads = { "IGST", "CGST", "SGST", "Cess" };
            int order = 1;

            foreach (var head in taxHeads)
            {
                var r3B = new AnnualSheetRow(12)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = $"ITC as per GSTR-3B ({head})",
                    Description = head,
                    SortOrder = order++,
                    IsSectionSummaryRow = true
                };

                var rInw = new AnnualSheetRow(12)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = $"ITC as per {targetReturnName} ({head})",
                    Description = head,
                    SortOrder = order++,
                    IsSectionSummaryRow = true
                };

                var rDiff = new AnnualSheetRow(12)
                {
                    Gstin = gstin,
                    GstinDisplay = gstinDisplay,
                    Section = $"Difference (GSTR-3B minus {targetReturnName})",
                    Description = head,
                    SortOrder = order++,
                    IsDifferenceRow = true
                };

                var inwSummary = rowsInward.FirstOrDefault(r =>
                    r.Section == "Total Net ITC (All Eligible Sections)" && r.Description == head);

                for (int m = 0; m < 12; m++)
                {
                    decimal val3B = Get3BVal(rows3B, "4(C) Net ITC Available (4A - 4B)", head, m, out bool has3B);
                    bool hasInw = inwSummary != null && m < inwSummary.HasMonthData.Length && inwSummary.HasMonthData[m];
                    decimal valInw = inwSummary != null && m < inwSummary.MonthlyValues.Length ? inwSummary.MonthlyValues[m] : 0m;

                    bool anyData = has3B || hasInw;
                    r3B.MonthlyValues[m] = val3B;
                    r3B.HasMonthData[m] = anyData;

                    rInw.MonthlyValues[m] = valInw;
                    rInw.HasMonthData[m] = anyData;

                    rDiff.MonthlyValues[m] = val3B - valInw;
                    rDiff.HasMonthData[m] = anyData;
                }

                result.Add(r3B);
                result.Add(rInw);
                result.Add(rDiff);

                if (Math.Abs(rDiff.TotalValue) >= 1m)
                {
                    alerts.Add(new AnnualReconciliationAlert
                    {
                        GstinDisplay = gstinDisplay,
                        ComparisonSheet = $"GSTR-3B vs {targetReturnName}",
                        MetricName = $"Net ITC ({head}) Difference",
                        DifferenceAmount = rDiff.TotalValue
                    });
                }
            }

            return result;
        }

        private static AnnualOverviewGstinSummary BuildOverviewSummaryForGstin(
            string gstin,
            string gstinDisplay,
            List<AnnualScannedFile> r3BFiles,
            List<AnnualScannedFile> r1Files,
            List<AnnualScannedFile> r2AFiles,
            List<AnnualScannedFile> r2BFiles,
            List<AnnualSheetRow> rows3B,
            List<AnnualSheetRow> rows1,
            List<AnnualSheetRow> rows2A,
            List<AnnualSheetRow> rows2B)
        {
            decimal outTaxable3B = Sum3BRow(rows3B, "3.1.A Outward taxable supplies (excluding zero rated)", "Taxable")
                                 + Sum3BRow(rows3B, "3.1.B Outward taxable supplies (zero rated)", "Taxable")
                                 + Sum3BRow(rows3B, "3.1.C Other outward supplies (nil-rated, exempted)", "Taxable")
                                 + Sum3BRow(rows3B, "3.1.E Non-GST outward supplies", "Taxable");

            decimal outTax3B = Sum3BRow(rows3B, "3.1.A Outward taxable supplies (excluding zero rated)", "IGST")
                             + Sum3BRow(rows3B, "3.1.A Outward taxable supplies (excluding zero rated)", "CGST")
                             + Sum3BRow(rows3B, "3.1.A Outward taxable supplies (excluding zero rated)", "SGST")
                             + Sum3BRow(rows3B, "3.1.A Outward taxable supplies (excluding zero rated)", "Cess")
                             + Sum3BRow(rows3B, "3.1.B Outward taxable supplies (zero rated)", "IGST")
                             + Sum3BRow(rows3B, "3.1.B Outward taxable supplies (zero rated)", "Cess");

            decimal netItc3B = Sum3BRow(rows3B, "4(C) Net ITC Available (4A - 4B)", "IGST")
                             + Sum3BRow(rows3B, "4(C) Net ITC Available (4A - 4B)", "CGST")
                             + Sum3BRow(rows3B, "4(C) Net ITC Available (4A - 4B)", "SGST")
                             + Sum3BRow(rows3B, "4(C) Net ITC Available (4A - 4B)", "Cess");

            decimal cashPaid3B = Sum3BRow(rows3B, "6.1 Tax paid in cash", "IGST")
                               + Sum3BRow(rows3B, "6.1 Tax paid in cash", "CGST")
                               + Sum3BRow(rows3B, "6.1 Tax paid in cash", "SGST")
                               + Sum3BRow(rows3B, "6.1 Tax paid in cash", "Cess");

            decimal itcPaid3B = Sum3BRow(rows3B, "6.1 Total Paid through ITC", "IGST")
                              + Sum3BRow(rows3B, "6.1 Total Paid through ITC", "CGST")
                              + Sum3BRow(rows3B, "6.1 Total Paid through ITC", "SGST")
                              + Sum3BRow(rows3B, "6.1 Total Paid through ITC", "Cess");

            decimal outTaxable1 = rows1
                .Where(r => r.Section == "Total Outward Supplies" && r.Description == "Taxable")
                .Sum(r => r.TotalValue);

            decimal outTax1 = rows1
                .Where(r => r.Section == "Total Outward Supplies" && r.Description != "Taxable")
                .Sum(r => r.TotalValue);

            decimal itc2A = rows2A
                .Where(r => r.Section == "Total Net ITC (All Eligible Sections)" && r.Description != "Taxable")
                .Sum(r => r.TotalValue);

            decimal itc2B = rows2B
                .Where(r => r.Section == "Total Net ITC (All Eligible Sections)" && r.Description != "Taxable")
                .Sum(r => r.TotalValue);

            return new AnnualOverviewGstinSummary
            {
                Gstin = gstin,
                GstinDisplay = gstinDisplay,
                StateName = GstStateHelper.GetStateName(gstin),
                TotalActiveMonths = 12,
                R3BMonthsCount = r3BFiles.Select(f => f.SortKey).Distinct().Count(),
                R1MonthsCount = r1Files.Select(f => f.SortKey).Distinct().Count(),
                R2AMonthsCount = r2AFiles.Select(f => f.SortKey).Distinct().Count(),
                R2BMonthsCount = r2BFiles.Select(f => f.SortKey).Distinct().Count(),
                OutwardTaxable3B = outTaxable3B,
                OutwardTax3B = outTax3B,
                OutwardTaxable1 = outTaxable1,
                OutwardTax1 = outTax1,
                NetItc3B = netItc3B,
                Itc2A = itc2A,
                Itc2B = itc2B,
                CashPaid3B = cashPaid3B,
                ItcPaid3B = itcPaid3B
            };
        }

        #endregion

        #region Helpers

        private static decimal Get3BVal(List<AnnualSheetRow> rows, string section, string desc, int m, out bool hasData)
        {
            var r = rows.FirstOrDefault(x => x.Section == section && x.Description == desc);
            if (r != null && m < r.MonthlyValues.Length)
            {
                hasData = r.HasMonthData[m];
                return r.MonthlyValues[m];
            }
            hasData = false;
            return 0m;
        }

        private static decimal Sum3BRow(List<AnnualSheetRow> rows, string section, string desc)
        {
            var r = rows.FirstOrDefault(x => x.Section == section && x.Description == desc);
            return r != null ? r.TotalValue : 0m;
        }

        private static JsonElement UnwrapData(JsonElement root)
        {
            if (root.ValueKind == JsonValueKind.Object &&
                root.TryGetProperty("data", out var dataProp) &&
                dataProp.ValueKind == JsonValueKind.Object)
            {
                return dataProp;
            }
            return root;
        }

        private static void ExtractInvoiceItems(
            JsonElement inv,
            out decimal txval,
            out decimal iamt,
            out decimal camt,
            out decimal samt,
            out decimal csamt)
        {
            txval = 0m; iamt = 0m; camt = 0m; samt = 0m; csamt = 0m;

            if (inv.TryGetProperty("itms", out var itms) && itms.ValueKind == JsonValueKind.Array && itms.GetArrayLength() > 0)
            {
                foreach (var itm in itms.EnumerateArray())
                {
                    var det = itm.TryGetProperty("itm_det", out var idet) && idet.ValueKind == JsonValueKind.Object
                        ? idet
                        : itm;

                    txval += GetDec(det, "txval");
                    iamt += GetDec(det, "iamt") + GetDec(det, "igst");
                    camt += GetDec(det, "camt") + GetDec(det, "cgst");
                    samt += GetDec(det, "samt") + GetDec(det, "sgst");
                    csamt += GetDec(det, "csamt") + GetDec(det, "cess");
                }
            }

            if (txval == 0m && iamt == 0m && camt == 0m && samt == 0m && csamt == 0m)
            {
                txval = GetDec(inv, "txval");
                iamt = GetDec(inv, "iamt") + GetDec(inv, "igst");
                camt = GetDec(inv, "camt") + GetDec(inv, "cgst");
                samt = GetDec(inv, "samt") + GetDec(inv, "sgst");
                csamt = GetDec(inv, "csamt") + GetDec(inv, "cess");
            }
        }

        private static bool IsDateBefore(string dateStr, DateTime threshold)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) return false;
            string[] formats = { "dd-MM-yyyy", "dd/MM/yyyy", "yyyy-MM-dd", "d-M-yyyy" };
            if (DateTime.TryParseExact(dateStr.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                return dt < threshold;
            }
            return false;
        }

        private static string GetStr(JsonElement el, string prop)
        {
            if (el.ValueKind == JsonValueKind.Object &&
                el.TryGetProperty(prop, out var v) &&
                v.ValueKind == JsonValueKind.String)
            {
                return v.GetString() ?? string.Empty;
            }
            return string.Empty;
        }

        private static decimal GetDec(JsonElement el, string prop)
        {
            if (el.ValueKind != JsonValueKind.Object || !el.TryGetProperty(prop, out var v)) return 0m;
            if (v.ValueKind == JsonValueKind.Number && v.TryGetDecimal(out decimal d)) return d;
            if (v.ValueKind == JsonValueKind.String && decimal.TryParse(v.GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal sd)) return sd;
            return 0m;
        }

        private static decimal GetNestedDec(JsonElement el, string parentProp, string childProp)
        {
            if (el.ValueKind == JsonValueKind.Object &&
                el.TryGetProperty(parentProp, out var parent) &&
                parent.ValueKind == JsonValueKind.Object)
            {
                return GetDec(parent, childProp);
            }
            return 0m;
        }

        #endregion
    }
}
