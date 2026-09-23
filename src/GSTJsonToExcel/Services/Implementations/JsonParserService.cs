using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class JsonParserService : IJsonParserService
    {
        private readonly ILoggingService _logger;

        public JsonParserService(ILoggingService logger)
        {
            _logger = logger;
        }

        public async Task<JsonParseResult> ParseAndFlattenAsync(
            string jsonFilePath,
            IProgress<ConversionProgress>? progress = null,
            CancellationToken cancellationToken = default)
        {
            _logger.LogInfo($"Starting JSON parsing for file: {Path.GetFileName(jsonFilePath)}");
            progress?.Report(new ConversionProgress(10, "Reading JSON", "Reading JSON file into memory..."));

            var fileInfo = new FileInfo(jsonFilePath);
            var result = new JsonParseResult
            {
                SourceFileName = fileInfo.Name,
                SourceFilePath = fileInfo.FullName,
                SourceFileSizeBytes = fileInfo.Length
            };

            await using var stream = File.OpenRead(jsonFilePath);
            using var document = await JsonDocument.ParseAsync(stream, new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip
            }, cancellationToken);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ConversionProgress(25, "Indexing Data", "Analyzing structure and indexing all leaf elements..."));

            // 1. Index every leaf node across the entire JSON tree
            IndexAllLeaves(document.RootElement, "$", result);
            _logger.LogInfo($"Indexed {result.TotalLeafCount} leaf values from source JSON.");

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new ConversionProgress(40, "Extracting Sections", "Building structured relational worksheets..."));

            // 2. Extract structured tables
            var tables = ExtractTables(document.RootElement, result);
            foreach (var table in tables)
            {
                result.Tables.Add(table);
            }

            // 3. Create the comprehensive Master_Data_Index table for guaranteed zero data loss
            var auditIndexTable = CreateMasterIndexTable(result.AllLeafNodes);
            result.Tables.Add(auditIndexTable);

            // 4. Create Empty_Sections table if any empty arrays or objects were encountered
            if (result.EmptySections.Count > 0)
            {
                var emptyTable = CreateEmptySectionsTable(result.EmptySections);
                result.Tables.Add(emptyTable);
            }

            progress?.Report(new ConversionProgress(60, "Parsing Complete", $"Successfully processed {result.Tables.Count} sections and {result.TotalLeafCount} fields."));
            _logger.LogInfo($"JSON flattening completed: {result.Tables.Count} tables generated.");

            return result;
        }

        private void IndexAllLeaves(JsonElement element, string currentPath, JsonParseResult result)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    var objEnumerator = element.EnumerateObject();
                    bool hasProps = false;
                    foreach (var prop in objEnumerator)
                    {
                        hasProps = true;
                        string childPath = $"{currentPath}.{prop.Name}";
                        IndexAllLeaves(prop.Value, childPath, result);
                    }
                    if (!hasProps)
                    {
                        result.EmptySections.Add(currentPath);
                    }
                    break;

                case JsonValueKind.Array:
                    int arrayLength = element.GetArrayLength();
                    if (arrayLength == 0)
                    {
                        result.EmptySections.Add(currentPath);
                    }
                    else
                    {
                        int index = 0;
                        foreach (var item in element.EnumerateArray())
                        {
                            string itemPath = $"{currentPath}[{index}]";
                            IndexAllLeaves(item, itemPath, result);
                            index++;
                        }
                    }
                    break;

                default:
                    // Primitive leaf
                    string fieldName = ExtractFieldNameFromPath(currentPath);
                    var flatVal = JsonFlatValue.FromJsonElement(element, fieldName);
                    result.AllLeafNodes[currentPath] = flatVal;
                    break;
            }
        }

        private List<JsonFlatTable> ExtractTables(JsonElement root, JsonParseResult parseResult)
        {
            var tables = new List<JsonFlatTable>();

            if (root.ValueKind != JsonValueKind.Object)
            {
                // Handle root array
                if (root.ValueKind == JsonValueKind.Array)
                {
                    ProcessArray(root, "Root_Data", "$", null, tables);
                }
                return tables;
            }

            // Extract Root/Master level scalars into General_Info
            var generalInfoTable = new JsonFlatTable
            {
                SheetName = "General_Info",
                FullJsonPath = "$"
            };

            var rootScalarsRow = new JsonFlatRow { SourceJsonPath = "$" };
            bool hasRootScalars = false;
            var parentContext = new Dictionary<string, JsonFlatValue>(StringComparer.OrdinalIgnoreCase);

            foreach (var prop in root.EnumerateObject())
            {
                if (prop.Value.ValueKind != JsonValueKind.Object && prop.Value.ValueKind != JsonValueKind.Array)
                {
                    var val = JsonFlatValue.FromJsonElement(prop.Value, prop.Name);
                    rootScalarsRow.Set(prop.Name, val);
                    hasRootScalars = true;

                    // Keep key parent identifiers (e.g. gstin, ret_period, fp) to pass to child sections
                    if (JsonFlatValue.IsIdentifierField(prop.Name))
                    {
                        parentContext[$"parent_{prop.Name}"] = val;
                    }
                }
            }

            if (hasRootScalars)
            {
                generalInfoTable.AddRow(rootScalarsRow);
                tables.Add(generalInfoTable);
            }

            // Process all top-level sections
            foreach (var prop in root.EnumerateObject())
            {
                string sectionName = prop.Name;
                string currentPath = $"$.{sectionName}";

                if (prop.Value.ValueKind == JsonValueKind.Array)
                {
                    ProcessArray(prop.Value, sectionName, currentPath, parentContext, tables);
                }
                else if (prop.Value.ValueKind == JsonValueKind.Object)
                {
                    ProcessObjectSection(prop.Value, sectionName, currentPath, parentContext, tables);
                }
            }

            return tables;
        }

        private void ProcessObjectSection(
            JsonElement sectionElement,
            string sectionName,
            string currentPath,
            Dictionary<string, JsonFlatValue> parentContext,
            List<JsonFlatTable> tables)
        {
            // Check what this object contains:
            // 1) Does it contain child arrays? (e.g. itc_elg has itc_avl, itc_rev, itc_inelg; returnsDbCdredList has tax_pay, tax_paid.pd_by_cash)
            // 2) Does it contain child objects or scalars?

            var childArrays = new List<JsonProperty>();
            var childObjects = new List<JsonProperty>();
            var directScalars = new List<JsonProperty>();

            foreach (var prop in sectionElement.EnumerateObject())
            {
                if (prop.Value.ValueKind == JsonValueKind.Array)
                {
                    childArrays.Add(prop);
                }
                else if (prop.Value.ValueKind == JsonValueKind.Object)
                {
                    childObjects.Add(prop);
                }
                else
                {
                    directScalars.Add(prop);
                }
            }

            // Extract all child arrays into their own sheets
            foreach (var childArr in childArrays)
            {
                string arrSheetName = SanitizeSheetName($"{sectionName}_{childArr.Name}");
                string arrPath = $"{currentPath}.{childArr.Name}";
                ProcessArray(childArr.Value, arrSheetName, arrPath, parentContext, tables);
            }

            // If the section contains child objects or direct scalars, create a structured section sheet
            if (childObjects.Count > 0 || directScalars.Count > 0)
            {
                // Check if child objects are also nested containers with arrays (like returnsDbCdredList or tax_paid)
                bool allChildObjectsAreContainers = true;
                foreach (var childObj in childObjects)
                {
                    bool hasNestedArray = childObj.Value.EnumerateObject().Any(p => p.Value.ValueKind == JsonValueKind.Array);
                    if (!hasNestedArray)
                    {
                        allChildObjectsAreContainers = false;
                        break;
                    }
                }

                if (childObjects.Count > 0 && allChildObjectsAreContainers)
                {
                    // Recurse into child containers
                    foreach (var childObj in childObjects)
                    {
                        string subName = $"{sectionName}_{childObj.Name}";
                        string subPath = $"{currentPath}.{childObj.Name}";
                        ProcessObjectSection(childObj.Value, subName, subPath, parentContext, tables);
                    }

                    // Direct scalars in this container if any
                    if (directScalars.Count > 0)
                    {
                        var directTable = new JsonFlatTable
                        {
                            SheetName = SanitizeSheetName($"{sectionName}_info"),
                            FullJsonPath = currentPath
                        };
                        var dRow = new JsonFlatRow { SourceJsonPath = currentPath };
                        AddParentContext(dRow, parentContext);
                        foreach (var sc in directScalars)
                        {
                            dRow.Set(sc.Name, JsonFlatValue.FromJsonElement(sc.Value, sc.Name));
                        }
                        directTable.AddRow(dRow);
                        tables.Add(directTable);
                    }
                }
                else
                {
                    // Regular section with structured data (e.g., sup_details, intr_ltfee, eco_dtls, taxpayble_bal)
                    var sectionTable = new JsonFlatTable
                    {
                        SheetName = SanitizeSheetName(sectionName),
                        FullJsonPath = currentPath
                    };

                    // If child objects represent categories (e.g. sup_details -> osup_det, osup_zero, etc.)
                    if (childObjects.Count > 0)
                    {
                        foreach (var childObj in childObjects)
                        {
                            // Check if this child object has nested arrays inside it (e.g., tax_paid -> pd_by_cash[])
                            var nestedArrays = childObj.Value.EnumerateObject().Where(p => p.Value.ValueKind == JsonValueKind.Array).ToList();
                            foreach (var nArr in nestedArrays)
                            {
                                string nArrSheet = SanitizeSheetName($"{sectionName}_{childObj.Name}_{nArr.Name}");
                                string nArrPath = $"{currentPath}.{childObj.Name}.{nArr.Name}";
                                ProcessArray(nArr.Value, nArrSheet, nArrPath, parentContext, tables);
                            }

                            // If child object has scalar properties or nested scalar objects, flatten into a row
                            var nonArrayProps = childObj.Value.EnumerateObject().Where(p => p.Value.ValueKind != JsonValueKind.Array).ToList();
                            if (nonArrayProps.Count > 0)
                            {
                                var row = new JsonFlatRow { SourceJsonPath = $"{currentPath}.{childObj.Name}" };
                                AddParentContext(row, parentContext);
                                row.Set("section_category", JsonFlatValue.FromString(childObj.Name));

                                FlattenObjectIntoRow(childObj.Value, "", row);
                                sectionTable.AddRow(row);
                            }
                        }
                    }

                    // Direct scalars in this section if any
                    if (directScalars.Count > 0)
                    {
                        var scalarRow = new JsonFlatRow { SourceJsonPath = currentPath };
                        AddParentContext(scalarRow, parentContext);
                        scalarRow.Set("section_category", JsonFlatValue.FromString("general_details"));
                        foreach (var sc in directScalars)
                        {
                            scalarRow.Set(sc.Name, JsonFlatValue.FromJsonElement(sc.Value, sc.Name));
                        }
                        sectionTable.AddRow(scalarRow);
                    }

                    if (sectionTable.RowCount > 0)
                    {
                        tables.Add(sectionTable);
                    }
                }
            }
        }

        private void ProcessArray(
            JsonElement arrayElement,
            string sheetName,
            string currentPath,
            Dictionary<string, JsonFlatValue>? parentContext,
            List<JsonFlatTable> tables)
        {
            int count = arrayElement.GetArrayLength();
            if (count == 0) return;

            var table = new JsonFlatTable
            {
                SheetName = SanitizeSheetName(sheetName),
                FullJsonPath = currentPath
            };

            int rowIndex = 1;
            foreach (var item in arrayElement.EnumerateArray())
            {
                var row = new JsonFlatRow
                {
                    SourceJsonPath = $"{currentPath}[{rowIndex - 1}]"
                };

                AddParentContext(row, parentContext);
                row.Set("_row_id", JsonFlatValue.FromString($"R_{rowIndex}"));

                if (item.ValueKind == JsonValueKind.Object)
                {
                    // Build local parent context for nested child arrays (e.g. invoice items)
                    var currentParentContext = new Dictionary<string, JsonFlatValue>(StringComparer.OrdinalIgnoreCase);
                    if (parentContext != null)
                    {
                        foreach (var kvp in parentContext) currentParentContext[kvp.Key] = kvp.Value;
                    }

                    // Look for key identifiers in this row
                    foreach (var prop in item.EnumerateObject())
                    {
                        if (prop.Value.ValueKind != JsonValueKind.Object && prop.Value.ValueKind != JsonValueKind.Array)
                        {
                            if (JsonFlatValue.IsIdentifierField(prop.Name))
                            {
                                currentParentContext[$"parent_{prop.Name}"] = JsonFlatValue.FromJsonElement(prop.Value, prop.Name);
                            }
                        }
                    }
                    currentParentContext["_parent_row_id"] = JsonFlatValue.FromString($"R_{rowIndex}");

                    // Check for nested child arrays
                    foreach (var prop in item.EnumerateObject())
                    {
                        if (prop.Value.ValueKind == JsonValueKind.Array)
                        {
                            string childSheetName = SanitizeSheetName($"{sheetName}_{prop.Name}");
                            string childPath = $"{currentPath}[{rowIndex - 1}].{prop.Name}";
                            ProcessArray(prop.Value, childSheetName, childPath, currentParentContext, tables);
                        }
                    }

                    FlattenObjectIntoRow(item, "", row);
                }
                else
                {
                    // Array of primitives
                    row.Set("value", JsonFlatValue.FromJsonElement(item, "value"));
                }

                table.AddRow(row);
                rowIndex++;
            }

            if (table.RowCount > 0)
            {
                tables.Add(table);
            }
        }

        private void FlattenObjectIntoRow(JsonElement objElement, string prefix, JsonFlatRow row)
        {
            foreach (var prop in objElement.EnumerateObject())
            {
                string colName = string.IsNullOrEmpty(prefix) ? prop.Name : $"{prefix}.{prop.Name}";

                if (prop.Value.ValueKind == JsonValueKind.Object)
                {
                    // Recurse into nested object
                    FlattenObjectIntoRow(prop.Value, colName, row);
                }
                else if (prop.Value.ValueKind != JsonValueKind.Array)
                {
                    // Scalar leaf
                    row.Set(colName, JsonFlatValue.FromJsonElement(prop.Value, prop.Name));
                }
            }
        }

        private static void AddParentContext(JsonFlatRow row, Dictionary<string, JsonFlatValue>? parentContext)
        {
            if (parentContext == null) return;
            foreach (var kvp in parentContext)
            {
                row.Set(kvp.Key, kvp.Value);
            }
        }

        private JsonFlatTable CreateMasterIndexTable(Dictionary<string, JsonFlatValue> allLeaves)
        {
            var table = new JsonFlatTable
            {
                SheetName = "All_Data_Index",
                FullJsonPath = "$"
            };

            int rowId = 1;
            foreach (var kvp in allLeaves)
            {
                var row = new JsonFlatRow { SourceJsonPath = kvp.Key };
                row.Set("Index", JsonFlatValue.FromString(rowId.ToString()));
                row.Set("JSON_Path", JsonFlatValue.FromString(kvp.Key));
                row.Set("Field_Name", JsonFlatValue.FromString(ExtractFieldNameFromPath(kvp.Key)));
                row.Set("Value", kvp.Value);
                row.Set("Data_Type", JsonFlatValue.FromString(kvp.Value.ValueKind.ToString()));
                table.AddRow(row);
                rowId++;
            }

            return table;
        }

        private JsonFlatTable CreateEmptySectionsTable(List<string> emptySections)
        {
            var table = new JsonFlatTable
            {
                SheetName = "Empty_Sections",
                FullJsonPath = "$"
            };

            int id = 1;
            foreach (var section in emptySections)
            {
                var row = new JsonFlatRow { SourceJsonPath = section };
                row.Set("Index", JsonFlatValue.FromString(id.ToString()));
                row.Set("Section_Path", JsonFlatValue.FromString(section));
                row.Set("Status", JsonFlatValue.FromString("Empty in Source JSON (0 records)"));
                table.AddRow(row);
                id++;
            }

            return table;
        }

        private static string ExtractFieldNameFromPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return string.Empty;

            int dotIndex = path.LastIndexOf('.');
            string candidate = dotIndex >= 0 ? path[(dotIndex + 1)..] : path;

            // Strip array brackets like "[0]" if present
            int bracketIndex = candidate.IndexOf('[');
            if (bracketIndex >= 0)
            {
                candidate = candidate[..bracketIndex];
            }

            return string.IsNullOrEmpty(candidate) ? path : candidate;
        }

        public static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Section";

            // Remove invalid characters for Excel sheet names: \ / ? * [ ] :
            string clean = Regex.Replace(name, @"[\\/?*\[\]:]", "_");

            // Max length in Excel is 31 characters
            if (clean.Length > 31)
            {
                // Take suffix or prefix
                clean = clean[^31..];
                if (clean.StartsWith("_")) clean = clean[1..];
            }

            return clean;
        }
    }
}
