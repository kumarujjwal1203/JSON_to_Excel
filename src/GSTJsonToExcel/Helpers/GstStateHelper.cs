using System.Collections.Generic;

namespace GSTJsonToExcel.Helpers
{
    public static class GstStateHelper
    {
        private static readonly Dictionary<string, string> StateMap = new()
        {
            { "01", "Jammu and Kashmir" },
            { "02", "Himachal Pradesh" },
            { "03", "Punjab" },
            { "04", "Chandigarh" },
            { "05", "Uttarakhand" },
            { "06", "Haryana" },
            { "07", "Delhi" },
            { "08", "Rajasthan" },
            { "09", "Uttar Pradesh" },
            { "10", "Bihar" },
            { "11", "Sikkim" },
            { "12", "Arunachal Pradesh" },
            { "13", "Nagaland" },
            { "14", "Manipur" },
            { "15", "Mizoram" },
            { "16", "Tripura" },
            { "17", "Meghalaya" },
            { "18", "Assam" },
            { "19", "West Bengal" },
            { "20", "Jharkhand" },
            { "21", "Odisha" },
            { "22", "Chhattisgarh" },
            { "23", "Madhya Pradesh" },
            { "24", "Gujarat" },
            { "26", "Dadra and Nagar Haveli and Daman and Diu" },
            { "27", "Maharashtra" },
            { "29", "Karnataka" },
            { "30", "Goa" },
            { "31", "Lakshadweep" },
            { "32", "Kerala" },
            { "33", "Tamil Nadu" },
            { "34", "Puducherry" },
            { "35", "Andaman and Nicobar Islands" },
            { "36", "Telangana" },
            { "37", "Andhra Pradesh" },
            { "38", "Ladakh" },
            { "97", "Other Territory" },
            { "99", "Centre Jurisdiction" }
        };

        public static string GetStateName(string? gstinOrCode)
        {
            if (string.IsNullOrWhiteSpace(gstinOrCode)) return string.Empty;

            string clean = gstinOrCode.Trim();
            string code = clean.Length >= 2 ? clean[..2] : clean;

            return StateMap.TryGetValue(code, out string? state) ? state : string.Empty;
        }

        public static string FormatGstinWithState(string gstin)
        {
            if (string.IsNullOrWhiteSpace(gstin)) return string.Empty;
            string state = GetStateName(gstin);
            return string.IsNullOrEmpty(state) ? gstin : $"{gstin} ({state})";
        }
    }
}
