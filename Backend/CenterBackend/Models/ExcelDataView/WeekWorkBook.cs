using CenterBackend.Models.CalculateData;

namespace CenterBackend.Models.ExcelDataView
{

    public class WeekWorkBook : BaseSheet
    {
        public List<WorkSheet1> WorkSheet1 { get; set; } = [];
        public List<WorkSheet2> WorkSheet2 { get; set; } = [];
        public List<WorkSheet3> WorkSheet3 { get; set; } = [];
        public List<WorkSheet4> WorkSheet4 { get; set; } = [];
        public List<WorkSheet5> WorkSheet5 { get; set; } = [];
        public List<WorkSheet6> WorkSheet6 { get; set; } = [];
        public List<WorkSheet7> WorkSheet7 { get; set; } = [];
        public List<WorkSheet8> WorkSheet8 { get; set; } = [];
        public MaterialDataRangeCollection? WorkSheet9 { get; set; }
    }

    public class WorkSheet1
    {
        public DateTime TimePoint { get; set; }
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
        public decimal? Cell14 { get; set; }
        public decimal? Cell15 { get; set; }
        public decimal? Cell16 { get; set; }
        public decimal? Cell17 { get; set; }
        public decimal? Cell18 { get; set; }
        public decimal? Cell19 { get; set; }
        public decimal? Cell20 { get; set; }
        public decimal? Cell21 { get; set; }
        public decimal? Cell22 { get; set; }
        public decimal? Cell23 { get; set; }
        public decimal? Cell24 { get; set; }
        public decimal? Cell25 { get; set; }
        public decimal? Cell26 { get; set; }
        public decimal? Cell27 { get; set; }
        public decimal? Cell28 { get; set; }

    }

    public class WorkSheet2
    {
        public DateTime TimePoint { get; set; }
        public decimal? Cell1 { get; set; }
        public decimal? Cell2 { get; set; }
        public decimal? Cell3 { get; set; }
        public decimal? Cell4 { get; set; }
        public decimal? Cell5 { get; set; }
    }

    public class WorkSheet3
    {
        public DateTime TimePoint { get; set; }
        public decimal? Cell1 { get; set; }
        public decimal? Cell2 { get; set; }
        public decimal? Cell3 { get; set; }
        public decimal? Cell4 { get; set; }
        public decimal? Cell5 { get; set; }
        public decimal? Cell6 { get; set; }
        public decimal? Cell7 { get; set; }
        public decimal? Cell8 { get; set; }
    }

    public class WorkSheet4
    {
        public DateTime TimePoint { get; set; }
        public DailyProductionReport? Data { get; set; }

    }

    public class WorkSheet5
    {
        public DateTime TimePoint { get; set; }
        public decimal? Cell1 { get; set; }
        public decimal? Cell2 { get; set; }
        public decimal? Cell3 { get; set; }
        public decimal? Cell4 { get; set; }
        public decimal? Cell5 { get; set; }
    }

    public class WorkSheet6
    {
        public DateTime TimePoint { get; set; }
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

    public class WorkSheet7
    {
        public DateTime TimePoint { get; set; }
        public decimal? Cell1 { get; set; }
        public decimal? Cell2 { get; set; }
        public decimal? Cell3 { get; set; }
    }
    public class WorkSheet8
    {
        public DateTime TimePoint { get; set; }
        public decimal? Cell1 { get; set; }
        public decimal? Cell2 { get; set; }
        public decimal? Cell3 { get; set; }
        public decimal? Cell4 { get; set; }
        public decimal? Cell5 { get; set; }
        public decimal? Cell6 { get; set; }
        public decimal? Cell7 { get; set; }
    }
}
