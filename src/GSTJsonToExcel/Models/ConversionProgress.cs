namespace GSTJsonToExcel.Models
{
    public class ConversionProgress
    {
        public int Percentage { get; set; }
        public string Phase { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;

        public ConversionProgress(int percentage, string phase, string message)
        {
            Percentage = percentage;
            Phase = phase;
            Message = message;
        }
    }
}
