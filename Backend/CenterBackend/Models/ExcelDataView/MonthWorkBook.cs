namespace CenterBackend.Models.ExcelDataView
{
    public class MonthWorkBook : BaseSheet
    {
        public MonthAnalysis MonthAnalysis { get; set; } = new();
    }
    public class MonthAnalysis
    {
        public int TimePoint { get; set; }
        public decimal? Cell1 { get; set; }
        public decimal? Cell2 { get; set; }
        public decimal? Cell3 { get; set; }
        public decimal? Cell4 { get; set; }
        public decimal? Cell5 { get; set; }
        public decimal? Cell6 { get; set; }
        public decimal? Cell7 { get; set; }
        public decimal? Cell8 { get; set; }
        public decimal? Cell9 { get; set; }
        public decimal? Cell10 { get; set; }
        public decimal? Cell11 { get; set; }
        public decimal? Cell12 { get; set; }
        public decimal? Cell13 { get; set; }
    }
}
