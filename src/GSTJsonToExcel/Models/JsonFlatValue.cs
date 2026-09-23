using System.Text.Json;

namespace GSTJsonToExcel.Models
{
    /// <summary>
    /// Represents an individual leaf value extracted from JSON with exact type and precision preservation.
    /// </summary>
    public class JsonFlatValue
    {
        public string RawString { get; set; } = string.Empty;
        public JsonValueKind ValueKind { get; set; } = JsonValueKind.Undefined;
        public decimal? DecimalValue { get; set; }
        public long? LongValue { get; set; }
        public bool? BoolValue { get; set; }
        public bool IsPreservedText { get; set; }
        public bool IsNull => ValueKind == JsonValueKind.Null;

        public static JsonFlatValue FromJsonElement(JsonElement element, string fieldName)
        {
            var value = new JsonFlatValue { ValueKind = element.ValueKind };

            switch (element.ValueKind)
            {
                case JsonValueKind.String:
                    value.RawString = element.GetString() ?? string.Empty;
                    // Check if this string is a numeric string with leading zeros or a sensitive identifier
                    value.IsPreservedText = ShouldPreserveAsText(value.RawString, fieldName);
                    break;

                case JsonValueKind.Number:
                    string rawNum = element.GetRawText();
                    value.RawString = rawNum;

                    // Parse as decimal first to prevent IEEE 754 floating point loss
                    if (decimal.TryParse(rawNum, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal decVal))
                    {
                        value.DecimalValue = decVal;
                    }

                    // Also record integer if there is no decimal point
                    if (!rawNum.Contains('.') && long.TryParse(rawNum, out long longVal))
                    {
                        value.LongValue = longVal;
                    }

                    // If field name indicates an ID/Code that shouldn't have number formatting quirks
                    if (IsIdentifierField(fieldName))
                    {
                        value.IsPreservedText = true;
                    }
                    break;

                case JsonValueKind.True:
                    value.BoolValue = true;
                    value.RawString = "true";
                    break;

                case JsonValueKind.False:
                    value.BoolValue = false;
                    value.RawString = "false";
                    break;

                case JsonValueKind.Null:
                    value.RawString = "(null)";
                    break;

                default:
                    value.RawString = element.GetRawText();
                    break;
            }

            return value;
        }

        public static JsonFlatValue FromString(string text, bool preserveText = true)
        {
            return new JsonFlatValue
            {
                RawString = text,
                ValueKind = JsonValueKind.String,
                IsPreservedText = preserveText
            };
        }

        private static bool ShouldPreserveAsText(string val, string fieldName)
        {
            if (string.IsNullOrEmpty(val)) return false;

            // Numbers with leading zeros like "00001234", "01"
            if (val.Length > 1 && val[0] == '0' && val.All(char.IsDigit))
            {
                return true;
            }

            return IsIdentifierField(fieldName);
        }

        public static bool IsIdentifierField(string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName)) return false;

            string lower = fieldName.ToLowerInvariant();
            return lower.Contains("gstin") ||
                   lower.Contains("ctin") ||
                   lower.Contains("inum") ||
                   lower.Contains("inv") ||
                   lower.Contains("hsn") ||
                   lower.Contains("sac") ||
                   lower.Contains("pin") ||
                   lower.Contains("pincode") ||
                   lower.Contains("debit_id") ||
                   lower.Contains("liab_id") ||
                   lower.Contains("trancd") ||
                   lower.Contains("ret_period") ||
                   lower.Contains("fp") ||
                   lower.Contains("pos") ||
                   lower.Contains("ref") ||
                   lower.Contains("code") ||
                   lower.Contains("pan") ||
                   lower.Contains("id");
        }
    }
}
