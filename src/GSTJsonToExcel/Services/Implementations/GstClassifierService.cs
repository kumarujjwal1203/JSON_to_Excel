using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class GstClassifierService : IGstClassifierService
    {
        private readonly ILoggingService _logger;

        public GstClassifierService(ILoggingService logger)
        {
            _logger = logger;
        }

        public async Task<(GstFileType Type, string Reason)> ClassifyFileAsync(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return (GstFileType.Unknown, "File does not exist.");
            }

            string fileName = Path.GetFileName(filePath);

            // 1. First Priority: Filename pattern matching
            var (typeFromName, reasonFromName) = ClassifyByFileName(fileName);
            if (typeFromName != GstFileType.Unknown)
            {
                return (typeFromName, reasonFromName);
            }

            // 2. Second Priority: Deep JSON content inspection
            return await ClassifyByContentAsync(filePath);
        }

        public (GstFileType Type, string Reason) ClassifyByFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return (GstFileType.Unknown, "Empty file name.");
            }

            string clean = Path.GetFileNameWithoutExtension(fileName);

            // Pattern: R1 / GSTR1 / GSTR-1
            if (Regex.IsMatch(clean, @"(?i)(^|[^a-z0-9])(r1|gstr1|gstr-1)([^a-z0-9]|$)", RegexOptions.CultureInvariant))
            {
                return (GstFileType.R1, "Identified by filename pattern (R1 / GSTR-1)");
            }

            // Pattern: R3A / R3B / GSTR3A / GSTR3B
            if (Regex.IsMatch(clean, @"(?i)(^|[^a-z0-9])(r3a|gstr3a|gstr-3a|r3b|gstr3b|gstr-3b)([^a-z0-9]|$)", RegexOptions.CultureInvariant))
            {
                return (GstFileType.R3A, "Identified by filename pattern (R3A / R3B)");
            }

            // Pattern: R2A / GSTR2A
            if (Regex.IsMatch(clean, @"(?i)(^|[^a-z0-9])(r2a|gstr2a|gstr-2a)([^a-z0-9]|$)", RegexOptions.CultureInvariant))
            {
                return (GstFileType.R2A, "Identified by filename pattern (R2A / GSTR-2A)");
            }

            // Pattern: R2B / GSTR2B
            if (Regex.IsMatch(clean, @"(?i)(^|[^a-z0-9])(r2b|gstr2b|gstr-2b)([^a-z0-9]|$)", RegexOptions.CultureInvariant))
            {
                return (GstFileType.R2B, "Identified by filename pattern (R2B / GSTR-2B)");
            }

            return (GstFileType.Unknown, "Filename is ambiguous; requires content inspection.");
        }

        private async Task<(GstFileType Type, string Reason)> ClassifyByContentAsync(string filePath)
        {
            try
            {
                await using var stream = File.OpenRead(filePath);
                using var doc = await JsonDocument.ParseAsync(stream, new JsonDocumentOptions
                {
                    AllowTrailingCommas = true,
                    CommentHandling = JsonCommentHandling.Skip
                });

                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object)
                {
                    return (GstFileType.Unknown, "Root JSON element is not an object.");
                }

                var propNames = root.EnumerateObject().Select(p => p.Name.ToLowerInvariant()).ToHashSet();

                // Check for R3A / R3B signatures:
                // Typically contains sup_details, intr_ltfee, eco_dtls, taxpayble, returnsDbCdredList, or itc_elg
                if (propNames.Contains("sup_details") ||
                    propNames.Contains("taxpayble") ||
                    propNames.Contains("returnsdbcdredlist") ||
                    (propNames.Contains("itc_elg") && propNames.Contains("intr_ltfee")))
                {
                    return (GstFileType.R3A, "Identified by content inspection (found sup_details / taxpayble signature)");
                }

                // Check for R2B signature:
                // Typically contains itc_summ, itc_unavl, or data with itc_avl
                if (propNames.Contains("itc_summ") || propNames.Contains("itc_unavl") || propNames.Contains("itcunavl"))
                {
                    return (GstFileType.R2B, "Identified by content inspection (found itc_summ / itc_unavl signature)");
                }

                // Check for R1 signatures:
                // Contains b2b along with outward supply sections like b2cs, hsn, cdnr, exp, doc_issue, nil, gt
                if (propNames.Contains("b2b") &&
                    (propNames.Contains("b2cs") ||
                     propNames.Contains("hsn") ||
                     propNames.Contains("cdnr") ||
                     propNames.Contains("doc_issue") ||
                     propNames.Contains("cur_gt") ||
                     propNames.Contains("exp")))
                {
                    return (GstFileType.R1, "Identified by content inspection (found GSTR-1 outward supply sections)");
                }

                // Check for R2A signature:
                // Contains b2b or cdn with supplier auto-drafted indicators (cfs, cfs3b, fldtr1)
                if (propNames.Contains("b2ba") || propNames.Contains("cdna") ||
                    (propNames.Contains("b2b") && (propNames.Contains("cdn") || propNames.Contains("isd"))))
                {
                    return (GstFileType.R2A, "Identified by content inspection (found GSTR-2A inward auto-drafted sections)");
                }

                // If only b2b is present, inspect first item for R1 vs R2A markers
                if (propNames.Contains("b2b"))
                {
                    var b2bProp = root.GetProperty("b2b");
                    if (b2bProp.ValueKind == JsonValueKind.Array && b2bProp.GetArrayLength() > 0)
                    {
                        var firstItem = b2bProp[0];
                        if (firstItem.ValueKind == JsonValueKind.Object)
                        {
                            var itemProps = firstItem.EnumerateObject().Select(p => p.Name.ToLowerInvariant()).ToHashSet();
                            if (itemProps.Contains("cfs") || itemProps.Contains("cfs3b") || itemProps.Contains("fldtr1"))
                            {
                                return (GstFileType.R2A, "Identified by content inspection (found supplier filing status cfs/cfs3b in B2B)");
                            }
                        }
                    }
                    return (GstFileType.R1, "Identified as R1 (found B2B invoice array)");
                }

                // Could not confidently determine
                return (GstFileType.Unknown, "Ambiguous JSON structure without recognized GST return signature.");
            }
            catch (JsonException ex)
            {
                return (GstFileType.Unknown, $"Invalid JSON content: {ex.Message}");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Content classification failed for {Path.GetFileName(filePath)}: {ex.Message}");
                return (GstFileType.Unknown, $"Inspection error: {ex.Message}");
            }
        }
    }
}
