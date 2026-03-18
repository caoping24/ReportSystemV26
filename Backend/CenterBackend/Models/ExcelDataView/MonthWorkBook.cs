namespace CenterBackend.Models.ExcelDataView
{
    public class MonthWorkBook : BaseSheet
    {
        public MonthAnalysis MonthAnalysis { get; set; } = new();
    }
    public class MonthAnalysis
    {
        public int TimePoint { get; set; }
        public float? Cell1 { get; set; }
        public float? Cell2 { get; set; }
        public float? Cell3 { get; set; }
        public float? Cell4 { get; set; }
        public float? Cell5 { get; set; }
        public float? Cell6 { get; set; }
        public float? Cell7 { get; set; }
        public float? Cell8 { get; set; }
        public float? Cell9 { get; set; }
        public float? Cell10 { get; set; }
        public float? Cell11 { get; set; }
        public float? Cell12 { get; set; }
        public float? Cell13 { get; set; }
    }
}
