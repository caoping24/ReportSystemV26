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
        public MaterialDataWeeklyCollection? WorkSheet9 { get; set; }  
    }

    public class WorkSheet1
    {
        public DateTime TimePoint { get; set; }
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
        public float? Cell14 { get; set; }
        public float? Cell15 { get; set; }
        public float? Cell16 { get; set; }
        public float? Cell17 { get; set; }
        public float? Cell18 { get; set; }
        public float? Cell19 { get; set; }
        public float? Cell20 { get; set; }
        public float? Cell21 { get; set; }
        public float? Cell22 { get; set; }
        public float? Cell23 { get; set; }
        public float? Cell24 { get; set; }
        public float? Cell25 { get; set; }
        public float? Cell26 { get; set; }
        public float? Cell27 { get; set; }
        public float? Cell28 { get; set; }

    }

    public class WorkSheet2
    {
        public DateTime TimePoint { get; set; }
        public float? Cell1 { get; set; }
        public float? Cell2 { get; set; }
        public float? Cell3 { get; set; }
        public float? Cell4 { get; set; }
        public float? Cell5 { get; set; }
    }

    public class WorkSheet3
    {
        public DateTime TimePoint { get; set; }
        public float? Cell1 { get; set; }
        public float? Cell2 { get; set; }
        public float? Cell3 { get; set; }
        public float? Cell4 { get; set; }
        public float? Cell5 { get; set; }
        public float? Cell6 { get; set; }
        public float? Cell7 { get; set; }
        public float? Cell8 { get; set; }
    }

    public class WorkSheet4
    {
        public DateTime TimePoint { get; set; }
        public DailyProductionReport? Data { get; set; }

    }

    public class WorkSheet5
    {
        public DateTime TimePoint { get; set; }
        public float? Cell1 { get; set; }
        public float? Cell2 { get; set; }
        public float? Cell3 { get; set; }
        public float? Cell4 { get; set; }
        public float? Cell5 { get; set; }
    }

    public class WorkSheet6
    {
        public DateTime TimePoint { get; set; }
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

    public class WorkSheet7
    {
        public DateTime TimePoint { get; set; }
        public float? Cell1 { get; set; }
        public float? Cell2 { get; set; }
        public float? Cell3 { get; set; }
    }
    public class WorkSheet8
    {
        public DateTime TimePoint { get; set; }
        public float? Cell1 { get; set; }
        public float? Cell2 { get; set; }
        public float? Cell3 { get; set; }
        public float? Cell4 { get; set; }
        public float? Cell5 { get; set; }
        public float? Cell6 { get; set; }
        public float? Cell7 { get; set; }
    }
}
