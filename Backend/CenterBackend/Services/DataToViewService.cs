using CenterBackend.IServices;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace CenterBackend.Services
{
    public class DataToViewService(IReportRepository<SourceData> SourceData, IReportRepository<OperatorInputData> rperatorInputData) : IDataToViewService
    {

        private readonly IReportRepository<SourceData> _sourceData = SourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = rperatorInputData;

        public bool DayGetMapData(DayWorkBook DayWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            var startTime = DayWorkBook.ReportedTime.Date;
            DayWorkBook.DaySheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();
            DayWorkBook.NightSheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();

            if (sourceData == null || sourceData.Count == 0)
                return false;

            var baseDate = startTime.AddHours(8);
            var dataPart1 = SortDataByTime(sourceData, baseDate, 25);
            var dataPart2 = SortDataByTime(operatorInputData, baseDate, 25);

            List<SourceData> source1 = dataPart1.Take(13).ToList();
            List<OperatorInputData> source2 = dataPart2.Take(13).ToList();
            DayMoveData(DayWorkBook.DaySheet, source1, source2);//白班

            source1 = dataPart1.Skip(12).Take(13).ToList();
            source2 = dataPart2.Skip(12).Take(13).ToList();
            DayMoveData(DayWorkBook.NightSheet, source1, source2);//夜班

            return true;
        }

        public bool MonthGetMapData(MonthWorkBook MonthWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();

            return false;

        }

        public bool YearGetMapData(YearWorkBook YearWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();



            return false;

        }

        public bool WeekGetMapData(WeekWorkBook WeekWorkBook, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            // var filteredData = sourceDataList.Where(data => data.Type == "SpecificType").ToList();

            return false;

        }

        //根据时间排序数据-原始数据
        private static List<SourceData> SortDataByTime(List<SourceData> sourceData, DateTime baseDate, int maxCount)
        {
            baseDate = baseDate.Date.AddHours(8);//从8点开始
            var sortedList = new List<SourceData>(new SourceData[maxCount]);
            for (int i = 0; i < maxCount; i++)
            {
                DateTime intervalStart = baseDate.AddHours(i);
                DateTime intervalEnd = intervalStart.AddHours(1);
                var data = sourceData.FirstOrDefault(x => x.ReportedTime >= intervalStart && x.ReportedTime < intervalEnd);//匹配时间
                if ( data != null)
                {
                    sortedList[i] = data;
                }
            }
            return sortedList;
        }
        //根据时间排序数据-操作员输入数据
        private static List<OperatorInputData> SortDataByTime(List<OperatorInputData> OperatorInputData, DateTime baseDate, int maxCount)
        {
            baseDate = baseDate.Date.AddHours(8);//从8点开始
            var sortedList = new List<OperatorInputData>(new OperatorInputData[maxCount]);
            for (int i = 0; i < maxCount; i++)
            {
                DateTime intervalStart = baseDate.AddHours(i);
                DateTime intervalEnd = intervalStart.AddHours(1);
                var data = OperatorInputData.FirstOrDefault(x => x.ReportedTime >= intervalStart && x.ReportedTime < intervalEnd);//匹配时间
                if (data != null)
                {
                    sortedList[i] = data;
                }
            }
            return sortedList;
        }
        private static void DayMoveData(List<DayWorkSheet> DayWorkSheet, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            var target = DayWorkSheet;
            var source1 = sourceData;
            var source2 = operatorInputData;
            //原始数据的映射
            for (int i = 0; i < 13; i++)
            {
                if (source1 == null || source1[i] == null)
                    continue;
                target[i].Cell1 = source1[i].Cell1;
                target[i].Cell2 = source1[i].Cell2;
                target[i].Cell3 = source1[i].Cell3;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell4;
                    var prevData = source1[i - 1]?.Cell4;// 可能为 null
                    if (currentVal != null && prevData != null)
                        target[i].Cell4 = (currentVal - prevData) / 1000;
                }
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell5;
                    var prevData = source1[i - 1]?.Cell5;
                    if (currentVal != null && prevData != null)
                        target[i].Cell5 = (currentVal - prevData) / 1000;
                }
                target[i].Cell6 = source1[i].Cell6;
                target[i].Cell7 = source1[i].Cell7;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell8;
                    var prevData = source1[i - 1]?.Cell8;
                    if (currentVal != null && prevData != null)
                        target[i].Cell8 = (currentVal - prevData) / 1000;
                }
                target[i].Cell9 = source1[i].Cell9;
                target[i].Cell10 = source1[i].Cell10;
                target[i].Cell11 = source1[i].Cell11;
                target[i].Cell12 = source1[i].Cell12;
                target[i].Cell13 = source1[i].Cell13;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell14;
                    var prevData = source1[i - 1]?.Cell14;
                    if (currentVal != null && prevData != null)
                        target[i].Cell14 = (currentVal - prevData) / 1000;
                }
                target[i].Cell15 = source1[i].Cell15;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell16;
                    var prevData = source1[i - 1]?.Cell16;
                    if (currentVal != null && prevData != null)
                        target[i].Cell16 = (currentVal - prevData) / 1000;
                }
                target[i].Cell17 = source1[i].Cell17;
                target[i].Cell18 = source1[i].Cell18;
                target[i].Cell19 = source1[i].Cell19;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell20;
                    var prevData = source1[i - 1]?.Cell20;
                    if (currentVal != null && prevData != null)
                        target[i].Cell20 = (currentVal - prevData) / 1000;
                }
                target[i].Cell21 = source1[i].Cell21;

                target[i].Cell22 = source1[i].Cell22 < 2 ? source1[i].Cell22: -1;//摩尔比 小于2 直接出 
                if (i == 12)// 最后一个值
                    target[i].Cell23 = source1[i].Cell23;
                target[i].Cell24 = source1[i].Cell24;
                target[i].Cell25 = source1[i].Cell25;
                target[i].Cell26 = source1[i].Cell26;
                target[i].Cell27 = source1[i].Cell27;
                target[i].Cell28 = source1[i].Cell28;
                //人工检测数据
                //target[i].Cell29 = source1[i].Cell29;
                //target[i].Cell30 = source1[i].Cell30;
                //target[i].Cell31 = source1[i].Cell31;
                //target[i].Cell32 = source1[i].Cell32;
                //target[i].Cell33 = source1[i].Cell33;
                //target[i].Cell34 = source1[i].Cell34;
                //target[i].Cell35 = source1[i].Cell35;

                target[i].Cell36 = source1[i].Cell36;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell37;
                    var prevData = source1[i - 1]?.Cell37;
                    if (currentVal != null && prevData != null)
                        target[i].Cell37 = (currentVal - prevData) / 1000;
                }
                target[i].Cell38 = source1[i].Cell38;
                target[i].Cell39 = source1[i].Cell39;
                target[i].Cell40 = source1[i].Cell40;
                target[i].Cell41 = source1[i].Cell41;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell42;
                    var prevData = source1[i - 1]?.Cell42;
                    if (currentVal != null && prevData != null)
                        target[i].Cell42 = (currentVal - prevData) / 1000;
                }
                //预留数据
                //target[i].Cell43 = source1[i].Cell43;
                //target[i].Cell44 = source1[i].Cell44;
                //target[i].Cell45 = source1[i].Cell45;
                //target[i].Cell46 = source1[i].Cell46;
                //target[i].Cell47 = source1[i].Cell47;
                //target[i].Cell48 = source1[i].Cell48;
                //target[i].Cell49 = source1[i].Cell49;
                //target[i].Cell50 = source1[i].Cell50;

                target[i].Cell51 = source1[i].Cell51;
                target[i].Cell52 = source1[i].Cell52;
                target[i].Cell53 = source1[i].Cell53;
                target[i].Cell54 = source1[i].Cell54;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell55;
                    var prevData = source1[i - 1]?.Cell55;
                    if (currentVal != null && prevData != null)
                        target[i].Cell55 = (currentVal - prevData) / 1000;
                }
                //人工检测数据
                //target[i].Cell56 = source1[i].Cell56;
                //target[i].Cell57 = source1[i].Cell57;
                //target[i].Cell58 = source1[i].Cell58;
                //target[i].Cell59 = source1[i].Cell59;
                //target[i].Cell60 = source1[i].Cell60;

                target[i].Cell61 = source1[i].Cell61;
                target[i].Cell62 = source1[i].Cell62;
                target[i].Cell63 = source1[i].Cell63;
                target[i].Cell64 = source1[i].Cell64;
                target[i].Cell65 = source1[i].Cell65;
                target[i].Cell66 = source1[i].Cell66;
                target[i].Cell67 = source1[i].Cell67;
                target[i].Cell68 = source1[i].Cell68;
                target[i].Cell69 = source1[i].Cell69;
                target[i].Cell70 = source1[i].Cell70;
                target[i].Cell71 = source1[i].Cell71;
                target[i].Cell72 = source1[i].Cell72;
                target[i].Cell73 = source1[i].Cell73;
                target[i].Cell74 = source1[i].Cell74;
                target[i].Cell75 = source1[i].Cell75;
                target[i].Cell76 = source1[i].Cell76;
                target[i].Cell77 = source1[i].Cell77;
                target[i].Cell78 = source1[i].Cell78;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell79;
                    var prevData = source1[i - 1]?.Cell79;
                    if (currentVal != null && prevData != null)
                        target[i].Cell79 = (currentVal - prevData) / 1000;
                }
                target[i].Cell80 = source1[i].Cell80;
                target[i].Cell81 = source1[i].Cell81;
                //人工检测数据
                //target[i].Cell82 = source1[i].Cell82;
                //target[i].Cell83 = source1[i].Cell83;
                //target[i].Cell84 = source1[i].Cell84;
                //target[i].Cell85 = source1[i].Cell85;
                //target[i].Cell86 = source1[i].Cell86;
                //target[i].Cell87 = source1[i].Cell87;

                target[i].Cell88 = source1[i].Cell88;
                target[i].Cell89 = source1[i].Cell89;
                target[i].Cell90 = source1[i].Cell90;
                target[i].Cell91 = source1[i].Cell91;
                target[i].Cell92 = source1[i].Cell92;
                //预留数据
                //target[i].Cell93 = source1[i].Cell93;
                //target[i].Cell94 = source1[i].Cell94;
                //target[i].Cell95 = source1[i].Cell95;
                //target[i].Cell96 = source1[i].Cell96;
                //target[i].Cell97 = source1[i].Cell97;
                //target[i].Cell98 = source1[i].Cell98;
                //target[i].Cell99 = source1[i].Cell99;
                //target[i].Cell100 = source1[i].Cell100;
                target[i].Cell101 = source1[i].Cell101;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell102;
                    var prevData = source1[i - 1]?.Cell102;
                    if (currentVal != null && prevData != null)
                        target[i].Cell102 = (currentVal - prevData) / 1000;
                }
                target[i].Cell103 = source1[i].Cell103;
                target[i].Cell104 = source1[i].Cell104;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell105;
                    var prevData = source1[i - 1]?.Cell105;
                    if (currentVal != null && prevData != null)
                        target[i].Cell105 = (currentVal - prevData) / 1000;
                }
                target[i].Cell106 = source1[i].Cell106;
                target[i].Cell107 = source1[i].Cell107;
                target[i].Cell108 = source1[i].Cell108;
                target[i].Cell109 = source1[i].Cell109;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell110;
                    var prevData = source1[i - 1]?.Cell110;
                    if (currentVal != null && prevData != null)
                        target[i].Cell110 = (currentVal - prevData) / 1000;
                }
                target[i].Cell111 = source1[i].Cell111;
                target[i].Cell112 = source1[i].Cell112;
                target[i].Cell113 = source1[i].Cell113;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell114;
                    var prevData = source1[i - 1]?.Cell114;
                    if (currentVal != null && prevData != null)
                        target[i].Cell114 = (currentVal - prevData) / 1000;
                }
                target[i].Cell115 = source1[i].Cell115;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell116;
                    var prevData = source1[i - 1]?.Cell116;
                    if (currentVal != null && prevData != null)
                        target[i].Cell116 = (currentVal - prevData) / 1000;
                }
                target[i].Cell117 = source1[i].Cell117;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell118;
                    var prevData = source1[i - 1]?.Cell118;
                    if (currentVal != null && prevData != null)
                        target[i].Cell118 = (currentVal - prevData) / 1000;
                }
                target[i].Cell119 = source1[i].Cell119;
                target[i].Cell120 = source1[i].Cell120;
                target[i].Cell121 = source1[i].Cell121;
                target[i].Cell122 = source1[i].Cell122;
                target[i].Cell123 = source1[i].Cell123;
                target[i].Cell124 = source1[i].Cell124;
                target[i].Cell125 = source1[i].Cell125;
                target[i].Cell126 = source1[i].Cell126;
                target[i].Cell127 = source1[i].Cell127;
                target[i].Cell128 = source1[i].Cell128;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell129;
                    var prevData = source1[i - 1]?.Cell129;
                    if (currentVal != null && prevData != null)
                        target[i].Cell129 = (currentVal - prevData) / 1000;
                }
                target[i].Cell130 = source1[i].Cell130;
                target[i].Cell131 = source1[i].Cell131;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell132;
                    var prevData = source1[i - 1]?.Cell132;
                    if (currentVal != null && prevData != null)
                        target[i].Cell132 = (currentVal - prevData) / 1000;
                }
                target[i].Cell133 = source1[i].Cell133;
                target[i].Cell134 = source1[i].Cell134;
                //人工检测数据
                //target[i].Cell135 = source1[i].Cell135;
                //target[i].Cell136 = source1[i].Cell136;
                //target[i].Cell137 = source1[i].Cell137;
                //target[i].Cell138 = source1[i].Cell138;
                //target[i].Cell139 = source1[i].Cell139;
                //target[i].Cell140 = source1[i].Cell140;
                //target[i].Cell141 = source1[i].Cell141;
                //预留数据
                //target[i].Cell142 = source1[i].Cell142;
                //target[i].Cell143 = source1[i].Cell143;
                //target[i].Cell144 = source1[i].Cell144;
                //target[i].Cell145 = source1[i].Cell145;
                //target[i].Cell146 = source1[i].Cell146;
                //target[i].Cell147 = source1[i].Cell147;
                //target[i].Cell148 = source1[i].Cell148;
                //target[i].Cell149 = source1[i].Cell149;
                //target[i].Cell150 = source1[i].Cell150;
            }
            //人工输入数据的映射
            for (int i = 0; i < 13; i++) {
                if (source2 == null || source2[i] == null)
                    continue;
                //人工检测数据
                target[i].Cell29 = source2[i].Cell1;
                target[i].Cell30 = source2[i].Cell2;
                target[i].Cell31 = source2[i].Cell3;
                target[i].Cell32 = source2[i].Cell4;
                target[i].Cell33 = source2[i].Cell5;
                target[i].Cell34 = source2[i].Cell6;
                target[i].Cell35 = source2[i].Cell7;
                //人工检测数据
                target[i].Cell56 = source2[i].Cell11;
                target[i].Cell57 = source2[i].Cell12;
                target[i].Cell58 = source2[i].Cell13;
                target[i].Cell59 = source2[i].Cell14;
                target[i].Cell60 = source2[i].Cell15;
                //人工检测数据
                target[i].Cell82 = source2[i].Cell21;
                target[i].Cell83 = source2[i].Cell22;
                target[i].Cell84 = source2[i].Cell23;
                target[i].Cell85 = source2[i].Cell24;
                target[i].Cell86 = source2[i].Cell25;
                target[i].Cell87 = source2[i].Cell26;
                //人工检测数据
                target[i].Cell135 = source2[i].Cell131;
                target[i].Cell136 = source2[i].Cell132;
                target[i].Cell137 = source2[i].Cell133;
                target[i].Cell138 = source2[i].Cell134;
                target[i].Cell139 = source2[i].Cell135;
                target[i].Cell140 = source2[i].Cell136;
                target[i].Cell141 = source2[i].Cell137;
            }
        }

    }
}
