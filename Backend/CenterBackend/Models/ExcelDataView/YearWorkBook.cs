namespace CenterBackend.Models.ExcelDataView
{
    public class YearWorkBook : BaseSheet
    {
        public List<YearAnalysis> YearAnalysis { get; set; } = new();
    }
    public class YearAnalysis : WorkSheet1
    {
    }
}
