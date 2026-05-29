namespace CenterBackend.Models.ExcelDataView
{
    public class MonthWorkBook : BaseSheet
    {
        public List<MonthAnalysis> MonthAnalysis { get; set; } = new();
    }
    public class MonthAnalysis : WorkSheet1
    {

    }
}
