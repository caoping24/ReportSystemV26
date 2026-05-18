using AngleSharp.Dom;
using CenterBackend.IServices;
using CenterBackend.Models.CalculateData;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using Masuit.Tools;
using Masuit.Tools.Models;
using Microsoft.Identity.Client;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using Org.BouncyCastle.Asn1.X509;
using System.Collections;
using System.Security.Cryptography.Xml;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace CenterBackend.Services
{
    public class DataToViewService(IReportRepository<SourceData> sourceData,IReportRepository<OperatorInputData> operatorInputData) : IDataToViewService
    {
        private readonly IReportRepository<SourceData> _sourceData = sourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = operatorInputData;

        public async Task<bool> DayGetMapDataAsync(DayWorkBook dayWorkBook)
        {
            DateTime startTime = dayWorkBook.ReportedTime.Date.AddHours(8);
            DateTime endTime = startTime.AddHours(25);
            var sourceData = await _sourceData.GetByDateTimeRangeAsync(startTime, endTime);
            if (sourceData == null || sourceData.Count == 0) return false;//未查到数据
            var operatorInputData = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endTime);
            MoveDataShifts(dayWorkBook, sourceData, operatorInputData);
            MoveDataShiftsAnalysis(dayWorkBook, sourceData, operatorInputData);
            MoveDataDayAnalysis(dayWorkBook, sourceData, operatorInputData);
            return true;
        }
        public async Task<bool> MonthGetMapDataAsync(MonthWorkBook monthWorkBook)
        {

            return true;
        }
        public async Task<bool> YearGetMapDataAsync(YearWorkBook yearWorkBook)
        {

            return true;
        }
        public async Task<bool> WeekGetMapDataAsync(WeekWorkBook WeekWorkBook)
        {
            DateTime startDay = WeekWorkBook.ReportedTime.Date.AddHours(8);  //开始日期的 8:00 
            DateTime endDay = WeekWorkBook.ReportedTime.Date.AddDays(7).AddHours(1); //一周后的 9:00
            List<SourceData> sourceData = await _sourceData.GetByDateTimeRangeAsync(startDay, endDay);
            List<OperatorInputData> operatorInputData = await _operatorInputData.GetByDateTimeRangeAsync(startDay, endDay);
            WeekMoveDataSheet1Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet2Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet3Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet4Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet5Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet6Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet7Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet8Async(WeekWorkBook, sourceData, operatorInputData);
            return true;
        }
        /***********************数据处理***********************/
        private static void MoveDataShifts(DayWorkBook dayWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            dayWorkBook.DaySheet = Enumerable.Range(0, 13).Select(_ => new SingleShift()).ToList();
            dayWorkBook.NightSheet = Enumerable.Range(0, 13).Select(_ => new SingleShift()).ToList();

            var startTime = dayWorkBook.ReportedTime.Date.AddHours(8);
            var dataPart1 = SortDataByTime(sourceDatas, startTime, 25);//原始数据
            var dataPart2 = SortDataByTime(operatorInputDatas, startTime, 25);//人工录入数据

            List<SourceData> source1;
            List<OperatorInputData> source2;

            source1 = dataPart1.Take(13).ToList();
            source2 = dataPart2.Take(13).ToList();
            SingleShiftMoveData(dayWorkBook.DaySheet, source1, source2);//白班

            source1 = dataPart1.Skip(12).Take(13).ToList();
            source2 = dataPart2.Skip(12).Take(13).ToList();
            SingleShiftMoveData(dayWorkBook.NightSheet, source1, source2);//夜班
        }
        private static void SingleShiftMoveData(List<SingleShift> dayWorkSheet, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            var target = dayWorkSheet;
            var source1 = sourceDatas;
            var source2 = operatorInputDatas;
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
                        target[i].Cell14 = (currentVal - prevData);
                }
                target[i].Cell15 = source1[i].Cell15;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell16;
                    var prevData = source1[i - 1]?.Cell16;
                    if (currentVal != null && prevData != null)
                        target[i].Cell16 = (currentVal - prevData);
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

                target[i].Cell22 = source1[i].Cell22 < 2 ? source1[i].Cell22 : -1;//摩尔比 小于2 直接出 
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
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell107;
                    var prevData = source1[i - 1]?.Cell107;
                    if (currentVal != null && prevData != null)
                        target[i].Cell107 = (currentVal - prevData);
                }
                target[i].Cell108 = source1[i].Cell108;
                target[i].Cell109 = source1[i].Cell109;
                target[i].Cell110 = source1[i].Cell110;
                target[i].Cell111 = source1[i].Cell111;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell112;
                    var prevData = source1[i - 1]?.Cell112;
                    if (currentVal != null && prevData != null)
                        target[i].Cell112 = (currentVal - prevData);//蒸汽本身就是吨
                }
                target[i].Cell113 = source1[i].Cell113;
                target[i].Cell114 = source1[i].Cell114;
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
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell120;
                    var prevData = source1[i - 1]?.Cell120;
                    if (currentVal != null && prevData != null)
                        target[i].Cell120 = (currentVal - prevData) / 1000;
                }
                target[i].Cell121 = source1[i].Cell121;
                target[i].Cell122 = source1[i].Cell122;
                target[i].Cell123 = source1[i].Cell123;
                target[i].Cell124 = source1[i].Cell124;
                target[i].Cell125 = source1[i].Cell125;
                target[i].Cell126 = source1[i].Cell126;
                target[i].Cell127 = source1[i].Cell127;
                target[i].Cell128 = source1[i].Cell128;
                target[i].Cell129 = source1[i].Cell129;
                target[i].Cell130 = source1[i].Cell130;
                //target[i].Cell131人工录入
                target[i].Cell132 = source1[i].Cell132;
                target[i].Cell133 = source1[i].Cell133;
                if (i != 0)// 每小时的差值
                {
                    var currentVal = source1[i].Cell134;
                    var prevData = source1[i - 1]?.Cell134;
                    if (currentVal != null && prevData != null)
                        target[i].Cell134 = (currentVal - prevData) / 1000;
                }
                target[i].Cell135 = source1[i].Cell135;
                target[i].Cell136 = source1[i].Cell136;
                //人工检测数据
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
            for (int i = 0; i < 13; i++)
            {
                if (source2 == null || source2[i] == null)
                    continue;
                //表2
                target[i].Cell29 = source2[i].Cell11;
                target[i].Cell30 = source2[i].Cell12;
                target[i].Cell31 = source2[i].Cell13;
                target[i].Cell32 = source2[i].Cell14;
                target[i].Cell33 = source2[i].Cell15;
                target[i].Cell34 = source2[i].Cell16;
                target[i].Cell35 = source2[i].Cell17;
                //表1
                target[i].Cell56 = source2[i].Cell1;
                target[i].Cell57 = source2[i].Cell2;
                target[i].Cell58 = source2[i].Cell3;
                target[i].Cell59 = source2[i].Cell4;
                target[i].Cell60 = source2[i].Cell5;
                //表4
                target[i].Cell82 = source2[i].Cell26;
                target[i].Cell83 = source2[i].Cell41;
                target[i].Cell84 = source2[i].Cell42;
                target[i].Cell85 = source2[i].Cell43;
                target[i].Cell86 = source2[i].Cell44;
                target[i].Cell87 = source2[i].Cell45;
                //表5
                target[i].Cell131 = source2[i].Cell36;
                target[i].Cell137 = source2[i].Cell56;
                target[i].Cell138 = source2[i].Cell62;

            }
        }
        private static void MoveDataShiftsAnalysis(DayWorkBook dayWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            var target = dayWorkBook.ShiftsAnalysis;
            DateTime startTime = dayWorkBook.ReportedTime.Date.AddHours(8);
            DateTime endTime = startTime.AddHours(25);
            var dayData = new DailyProductionReport(startTime, sourceDatas, operatorInputDatas);
            target.Data = dayData;
            target.TimePoint = startTime;
        }
        private static void MoveDataDayAnalysis(DayWorkBook dayWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            DayAnalysis target = dayWorkBook.DayAnalysis;

            DateTime startTime = dayWorkBook.ReportedTime.Date.AddHours(8);
            DateTime endTime = startTime.AddHours(25);

            var dayData = new DailyProductionReport(startTime, sourceDatas, operatorInputDatas);
            var dayYield = dayData?.Cell2 ?? 0;
            var Materialcollection = new MaterialDailyCollection(startTime, dayYield, sourceDatas, operatorInputDatas);

            target.Cell1 = dayData?.Cell1??0; //计算本日收率
            
            target.Cell3 = Materialcollection.MaterialDatas[1].Specific;     //氨单耗
            target.Cell4 = Materialcollection.MaterialDatas[2].Specific;     //稀硫酸单耗
            target.Cell5 = Materialcollection.MaterialDatas[3].Specific;     //羟基浓度 

            target.Cell2 = Materialcollection.MaterialDatas[0].Specific * target.Cell5 ;   //羟基单耗 配后累计(l)*配后浓度(g/l)

            target.Cell6 = Materialcollection.MaterialDatas[4].Specific;     //氨腈摩尔比
            target.Cell7 = MathTools.ResidenceTimeSeconds(15, 22, 300);  // 反应时间 固定值
            target.Cell8 = Materialcollection.MaterialDatas[6].Specific;     // 反应压力
            target.Cell9 = Materialcollection.MaterialDatas[7].Specific;     //羟基加热温度
            target.Cell10 = Materialcollection.MaterialDatas[8].Specific;     //氨汽混合温度
            target.Cell11 = Materialcollection.MaterialDatas[9].Specific;        //管反热点温度
            target.Cell12 = Materialcollection.MaterialDatas[10].Specific;        //预冷器结晶温度
            target.Cell13 = Materialcollection.MaterialDatas[11].Specific;       //一次结晶温度
            target.Cell14 = Materialcollection.MaterialDatas[12].Specific;       //降膜蒸发温度
            target.Cell15 = Materialcollection.MaterialDatas[13].Specific;       //二次结晶温度
                                                                            //Materialcollection.WeeklyCollections[14]        //脱盐水
            target.Cell21 = Materialcollection.MaterialDatas[15].Specific;       //废液排放

            if (operatorInputDatas != null)
            {
                var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                target.Cell16 = CalculateAverage(operatorInputData, x => x.Cell11);  //	二乙腈含量
                target.Cell17 = CalculateAverage(operatorInputData, x => x.Cell13);  //	羟基乙腈残余
                target.Cell18 = CalculateAverage(operatorInputData, x => x.Cell17);  //	pH值
                target.Cell19 = CalculateAverage(operatorInputData, x => x.Cell15);  //	甘氨腈
                target.Cell20 = CalculateAverage(operatorInputData, x => x.Cell16);  //	三乙腈

                target.Cell22 = CalculateAverage(operatorInputData, x => x.Cell74);  //	羟基乙腈
                target.Cell23 = CalculateAverage(operatorInputData, x => x.Cell75);  //	硫铵
                target.Cell24 = CalculateAverage(operatorInputData, x => x.Cell76);  //	二乙腈
                target.Cell25 = CalculateAverage(operatorInputData, x => x.Cell77);  //	甘氨腈
                target.Cell26 = CalculateAverage(operatorInputData, x => x.Cell78);  //	三乙腈
                target.Cell27 = CalculateAverage(operatorInputData, x => x.Cell79); //	其它
                target.Cell28 = CalculateAverage(operatorInputData, x => x.Cell80);	//	水分
            }
        }
        private static void MoveDataMonthAnalysis(MonthWorkBook monthWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas, List<OperatorInputData> operatorInputDatasLastMonth)
        {

        }
        private static void MoveDataYearAnalysis(YearWorkBook yearWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas, List<OperatorInputData> operatorInputDatasLastYear)
        {

        }
        private static bool WeekMoveDataSheet1Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet1 = Enumerable.Range(1, 1).Select(_ => new WorkSheet1()).ToList();
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            startTime = currentWeekFirstDay;
            endTime = startTime.AddDays(7).AddHours(1);
            var data = new MaterialDailyCollection(startTime, 2104.62f, sourceDatas, operatorInputDatas);
            var WeekRangeData = CalculateDailyProductionReportRange(startTime, endTime, sourceDatas, operatorInputDatas);
            List<float?> WeekRangeYield = WeekRangeData == null
                            ? new List<float?>() // 集合为null时返回空列表，避免空引用异常
                            : WeekRangeData.Select(report => report?.Cell2).ToList();
            var Materialcollection = new MaterialDataWeeklyCollection(startTime, WeekRangeYield, sourceDatas, operatorInputDatas);
            int i = 0;
            WeekWorkBook.WorkSheet1[i].Cell1 = CalculateWeighted(WeekRangeData, x => x.Cell1, x => x.Cell3); //计算本周收率
            WeekWorkBook.WorkSheet1[i].Cell2 = Materialcollection.WeeklyCollections[0];     //羟基单耗
            WeekWorkBook.WorkSheet1[i].Cell3 = Materialcollection.WeeklyCollections[1];     //氨单耗
            WeekWorkBook.WorkSheet1[i].Cell4 = Materialcollection.WeeklyCollections[2];     //稀硫酸单耗
            WeekWorkBook.WorkSheet1[i].Cell5 = Materialcollection.WeeklyCollections[3];     //羟基浓度 
            WeekWorkBook.WorkSheet1[i].Cell6 = Materialcollection.WeeklyCollections[4];     //氨腈摩尔比
            WeekWorkBook.WorkSheet1[i].Cell7 = MathTools.ResidenceTimeSeconds(15, 22, 300);  // 反应时间 固定值
            WeekWorkBook.WorkSheet1[i].Cell8 = Materialcollection.WeeklyCollections[6];     // 反应压力
            WeekWorkBook.WorkSheet1[i].Cell9 = Materialcollection.WeeklyCollections[7];     //羟基加热温度
            WeekWorkBook.WorkSheet1[i].Cell10 = Materialcollection.WeeklyCollections[8];     //氨汽混合温度
            WeekWorkBook.WorkSheet1[i].Cell11 = Materialcollection.WeeklyCollections[9];        //管反热点温度
            WeekWorkBook.WorkSheet1[i].Cell12 = Materialcollection.WeeklyCollections[10];        //预冷器结晶温度
            WeekWorkBook.WorkSheet1[i].Cell13 = Materialcollection.WeeklyCollections[11];       //一次结晶温度
            WeekWorkBook.WorkSheet1[i].Cell14 = Materialcollection.WeeklyCollections[12];       //降膜蒸发温度
            WeekWorkBook.WorkSheet1[i].Cell15 = Materialcollection.WeeklyCollections[13];       //二次结晶温度
                                              //Materialcollection.WeeklyCollections[14]        //脱盐水
            WeekWorkBook.WorkSheet1[i].Cell21 = Materialcollection.WeeklyCollections[15];       //废液排放
            if (operatorInputDatas != null)
            {
                var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                WeekWorkBook.WorkSheet1[i].Cell16 = CalculateAverage(operatorInputData, x => x.Cell11);  //	二乙腈含量
                WeekWorkBook.WorkSheet1[i].Cell17 = CalculateAverage(operatorInputData, x => x.Cell13);  //	羟基乙腈残余
                WeekWorkBook.WorkSheet1[i].Cell18 = CalculateAverage(operatorInputData, x => x.Cell17);  //	pH值
                WeekWorkBook.WorkSheet1[i].Cell19 = CalculateAverage(operatorInputData, x => x.Cell15);  //	甘氨腈
                WeekWorkBook.WorkSheet1[i].Cell20 = CalculateAverage(operatorInputData, x => x.Cell16);  //	三乙腈

                WeekWorkBook.WorkSheet1[i].Cell22 = CalculateAverage(operatorInputData, x => x.Cell74);  //	羟基乙腈
                WeekWorkBook.WorkSheet1[i].Cell23 = CalculateAverage(operatorInputData, x => x.Cell75);  //	硫铵
                WeekWorkBook.WorkSheet1[i].Cell24 = CalculateAverage(operatorInputData, x => x.Cell76);  //	二乙腈
                WeekWorkBook.WorkSheet1[i].Cell25 = CalculateAverage(operatorInputData, x => x.Cell77);  //	甘氨腈
                WeekWorkBook.WorkSheet1[i].Cell26 = CalculateAverage(operatorInputData, x => x.Cell78);  //	三乙腈
                WeekWorkBook.WorkSheet1[i].Cell27 = CalculateAverage(operatorInputData, x => x.Cell79); //	其它
                WeekWorkBook.WorkSheet1[i].Cell28 = CalculateAverage(operatorInputData, x => x.Cell80);	//	水分
            }
            WeekWorkBook.WorkSheet9 = Materialcollection;
            return true;
        }
        private static bool WeekMoveDataSheet2Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet2 = Enumerable.Range(1, 7).Select(_ => new WorkSheet2()).ToList();
            var datalist = WeekWorkBook.WorkSheet2;
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            for (var i = 0; i < 7; i++)
            {
                startTime = currentWeekFirstDay.AddDays(i);
                endTime = startTime.AddDays(1);
                var dataItem = datalist[i];
                dataItem.TimePoint = startTime;
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (operatorInputData.Count == 0)
                        continue;
                    
                    dataItem.Cell1 = CalculateAverage(operatorInputData, x => x.Cell1);
                    dataItem.Cell2 = CalculateAverage(operatorInputData, x => x.Cell2);
                    dataItem.Cell3 = CalculateAverage(operatorInputData, x => x.Cell3);
                    dataItem.Cell4 = CalculateAverage(operatorInputData, x => x.Cell4);
                    dataItem.Cell5 = CalculateAverage(operatorInputData, x => x.Cell5);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet3Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet3 = Enumerable.Range(1, 7).Select(_ => new WorkSheet3()).ToList();
            var datalist = WeekWorkBook.WorkSheet3;
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            for (var i = 0; i < 7; i++)
            {
                startTime = currentWeekFirstDay.AddDays(i);
                endTime = startTime.AddDays(1);
                var dataItem = datalist[i];
                dataItem.TimePoint = startTime;
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (operatorInputData.Count == 0)
                        continue;
                    dataItem.Cell1 = CalculateAverage(operatorInputData, x => x.Cell11);
                    dataItem.Cell2 = CalculateAverage(operatorInputData, x => x.Cell12);
                    dataItem.Cell3 = CalculateAverage(operatorInputData, x => x.Cell13);
                    dataItem.Cell4 = CalculateAverage(operatorInputData, x => x.Cell14);
                    dataItem.Cell5 = CalculateAverage(operatorInputData, x => x.Cell15);
                    dataItem.Cell6 = CalculateAverage(operatorInputData, x => x.Cell16);
                    dataItem.Cell7 = CalculateAverage(operatorInputData, x => x.Cell17);
                    dataItem.Cell8 = CalculateAverage(operatorInputData, x => x.Cell18);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet4Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet4 = Enumerable.Range(1, 7).Select(_ => new WorkSheet4()).ToList();

            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            DateTime startTime = currentWeekFirstDay;
            for (var i = 0; i < 7; i++)
            {
                
                var dataItem = WeekWorkBook.WorkSheet4[i];
                if (sourceDatas != null && operatorInputDatas != null)
                {
                    var dayData = new DailyProductionReport(startTime, sourceDatas, operatorInputDatas);
                    dataItem.Data= dayData;
                    dataItem.TimePoint = startTime;
                }
                startTime = startTime.AddDays(1);
            }
            return true;
        }
        private static bool WeekMoveDataSheet5Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet5 = Enumerable.Range(1, 7).Select(_ => new WorkSheet5()).ToList();
            var datalist = WeekWorkBook.WorkSheet5;
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 7; i++)
            {
                startTime = currentWeekFirstDay.AddDays(i);
                endTime = startTime.AddDays(1);
                var dataItem = datalist[i];
                dataItem.TimePoint = startTime;
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (operatorInputData.Count == 0)
                        continue;
                    
                    dataItem.Cell1 = CalculateAverage(operatorInputData, x => x.Cell41);
                    dataItem.Cell2 = CalculateAverage(operatorInputData, x => x.Cell42);
                    dataItem.Cell3 = CalculateAverage(operatorInputData, x => x.Cell43);
                    dataItem.Cell4 = CalculateAverage(operatorInputData, x => x.Cell44);
                    dataItem.Cell5 = CalculateAverage(operatorInputData, x => x.Cell45);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet6Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet6 = Enumerable.Range(1, 7).Select(_ => new WorkSheet6()).ToList();
            var datalist = WeekWorkBook.WorkSheet6;
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 7; i++)
            {
                startTime = currentWeekFirstDay.AddDays(i);
                endTime = startTime.AddDays(1);
                var dataItem = datalist[i];
                dataItem.TimePoint = startTime;
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (operatorInputData.Count == 0)
                        continue;
                    
                    dataItem.Cell1 = CalculateAverage(operatorInputData, x => x.Cell51);
                    dataItem.Cell2 = CalculateAverage(operatorInputData, x => x.Cell52);
                    dataItem.Cell3 = CalculateAverage(operatorInputData, x => x.Cell53);
                    dataItem.Cell4 = CalculateAverage(operatorInputData, x => x.Cell54);
                    dataItem.Cell5 = CalculateAverage(operatorInputData, x => x.Cell55);
                    dataItem.Cell6 = CalculateAverage(operatorInputData, x => x.Cell56);
                    dataItem.Cell7 = CalculateAverage(operatorInputData, x => x.Cell57);
                    dataItem.Cell8 = CalculateAverage(operatorInputData, x => x.Cell58);
                    dataItem.Cell9 = CalculateAverage(operatorInputData, x => x.Cell59);
                    dataItem.Cell10 = CalculateAverage(operatorInputData, x => x.Cell60);
                    dataItem.Cell11 = CalculateAverage(operatorInputData, x => x.Cell61);
                    dataItem.Cell12 = CalculateAverage(operatorInputData, x => x.Cell62);
                    dataItem.Cell13 = CalculateAverage(operatorInputData, x => x.Cell63);
                }
            }
            return true;

        }
        private static bool WeekMoveDataSheet7Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet7 = Enumerable.Range(1, 7).Select(_ => new WorkSheet7()).ToList();
            var datalist = WeekWorkBook.WorkSheet7;
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            for (var i = 0; i < 7; i++)
            {
                startTime = currentWeekFirstDay.AddDays(i);
                endTime = startTime.AddDays(1);
                var dataItem = datalist[i];
                dataItem.TimePoint = startTime;
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (operatorInputData.Count == 0)
                        continue;
                    
                    dataItem.Cell1 = CalculateAverage(operatorInputData, x => x.Cell71);
                    dataItem.Cell2 = CalculateAverage(operatorInputData, x => x.Cell72);
                    dataItem.Cell3 = CalculateAverage(operatorInputData, x => x.Cell73);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet8Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet8 = Enumerable.Range(1, 7).Select(_ => new WorkSheet8()).ToList();
            var datalist = WeekWorkBook.WorkSheet8;
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            for (var i = 0; i < 7; i++)
            {
                startTime = currentWeekFirstDay.AddDays(i);
                endTime = startTime.AddDays(1);
                var dataItem = datalist[i];
                dataItem.TimePoint = startTime;
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (operatorInputData.Count == 0)
                        continue;
                    dataItem.Cell1 = CalculateAverage(operatorInputData, x => x.Cell74);
                    dataItem.Cell2 = CalculateAverage(operatorInputData, x => x.Cell75);
                    dataItem.Cell3 = CalculateAverage(operatorInputData, x => x.Cell76);
                    dataItem.Cell4 = CalculateAverage(operatorInputData, x => x.Cell77);
                    dataItem.Cell5 = CalculateAverage(operatorInputData, x => x.Cell78);
                    dataItem.Cell6 = CalculateAverage(operatorInputData, x => x.Cell79);
                    dataItem.Cell7 = CalculateAverage(operatorInputData, x => x.Cell80);
                }
            }
            return true;
        }
        /***********************辅助方法***********************/
        /// <summary>
        /// 统计单个字段值落在 [minRange, maxRange] 范围内的数量(跳过null)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="selector">字段选择器(如 x => x.Cell26)</param>
        /// <param name="minRange">区间下限</param>
        /// <param name="maxRange">区间上限</param>
        public static (int nonNullTotal, int qualifiedCount) CountValueInRange<T>(IEnumerable<T> data, Func<T, float?> selector, float minRange, float maxRange)
        {
            if (data == null || !data.Any())//空数据校验
                return (0,0);
            var nonNullTotal = data        
                            .Select(selector)           
                            .Where(x => x != null)     
                            .Count();
            var qualifiedCount = data
                            .Select(selector)
                            .Where(x => x != null)
                            .Where(x => x >= minRange && x <= maxRange)
                            .Count();
            return (nonNullTotal, qualifiedCount);
        }
        /// <summary>
        /// 统计两个字段的比值落在 [minRange, maxRange] 范围内的数量(跳过null、避免除零)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="selector1">字段选择器 分子</param>
        /// <param name="selector2">字段选择器 分母</param>
        /// <param name="minRange">区间下限</param>
        /// <param name="maxRange">区间上限</param>
        public static (int nonNullTotal, int qualifiedCount) CountRatioInRange<T>(IEnumerable<T> data, Func<T, float?> selector1, Func<T, float?> selector2, float minRange, float maxRange)
        {
            if (data == null || !data.Any())//空数据校验
                return (0, 0);
            var nonNullTotal = data
                                .Where(x => selector1 != null && selector2 != null && selector2(x) != 0)
                                .Count();
            var qualifiedCount = data
                                .Where(x => selector1 != null && selector2 != null && selector2(x) != 0)
                                .Where(x => selector1(x)/selector2(x)>= minRange&& selector1(x) / selector2(x) <= maxRange )
                                .Count();
            return (nonNullTotal, qualifiedCount);
        }
        public static List<SourceData> SortDataByTime(List<SourceData> sourceData, DateTime baseDate, int maxCount)
        {
            baseDate = baseDate.Date.AddHours(8);//从8点开始
            var sortedList = new List<SourceData>(new SourceData[maxCount]);
            for (int i = 0; i < maxCount; i++)
            {
                DateTime intervalStart = baseDate.AddHours(i);
                DateTime intervalEnd = intervalStart.AddHours(1);
                var data = sourceData.FirstOrDefault(x => x.ReportedTime >= intervalStart && x.ReportedTime < intervalEnd);//匹配时间
                if (data != null)
                {
                    sortedList[i] = data;
                }
            }
            return sortedList;
        }
        //根据时间排序数据-操作员输入数据
        public static List<OperatorInputData> SortDataByTime(List<OperatorInputData> OperatorInputData, DateTime baseDate, int maxCount)
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
        public static DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = (int)dt.DayOfWeek - (int)DayOfWeek.Monday;
            if (diff < 0) diff += 7;
            return dt.AddDays(-diff).Date;
        }
        public static float? CalculateAverage<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())// 空数据校验
                return null;
            var validValues = data.Select(selector).OfType<float>();
            float sum = 0f;
            int count = 0;
            foreach (var value in validValues)
            {
                sum += value;
                count++;
            }
            return count > 0 ? sum / count : (float?)null;
        }
        public static float? CalculateFirstLastDifference<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())
                return null;

            var nonNullValues = data
                .Select(selector)
                .Where(x => x.HasValue)
                .Select(x => x.GetValueOrDefault())
                .ToList();
            if (nonNullValues.Count < 2)
                return null;
            float firstValue = nonNullValues.First();//计算差值
            float lastValue = nonNullValues.Last();
            float difference = lastValue - firstValue;
            return difference;
        }
        /// <param name="mode">mode：模式开关(true = 区间校验 [min,max]，false = 上限校验 [≤max])</param>
        public static float? CalculateQualifiedRate<T>(IEnumerable<T> data, Func<T, float?> selector, bool mode, float qualifiedValue, float qualifiedValuediff)
        {
            var maxQualifiedValue = qualifiedValue + qualifiedValuediff;
            var minQualifiedValue = qualifiedValue - qualifiedValuediff;
            if (data == null || !data.Any())// 空数据校验
                return null;
            var nonNullValues = data// 筛选非null的float值并计算平均值
                .Select(selector)
                .Where(x => x.HasValue)
                .Select(x => x.GetValueOrDefault())
                .ToList();
            if (nonNullValues.Count == 0)
                return null;

            int qualifiedCount = mode
                // mode=true：值在 [min, max] 区间内为合格
                ? nonNullValues.Count(value => value >= minQualifiedValue && value <= maxQualifiedValue)
                // mode=false：值 ≤ max 为合格
                : nonNullValues.Count(value => value <= maxQualifiedValue);
            float qualifiedRate = (qualifiedCount / (float)nonNullValues.Count) * 100f;// 计算合格利率(合格数/总有效数 * 100%)，保留3位小数
            return (float)Math.Round(qualifiedRate, 3);
        }
        public static float? CalculateSum<T>(IEnumerable<T> data, Func<T, float?> selector)//非null值的总和
        {
            if (data == null || !data.Any())//空数据校验
                return null;
            var nonNullValues = data        //筛选非null的float值
                .Select(selector)           // 提取float?字段
                .Where(x => x.HasValue)     // 过滤掉null值
                .Select(x => x.GetValueOrDefault())      // 转换为float(非可空)
                .ToList();
            return nonNullValues.Count != 0 ? nonNullValues.Sum() : (float?)null;//计算总和
        }
        /// <summary>
        /// 所有列求两列加权平均
        /// </summary>
        /// <param name="dataList"></param>
        /// <param name="getValue">含量</param>
        /// <param name="getWeight">总量</param>
        /// <returns></returns>
        public static float CalculateWeighted(
                                                List<DailyProductionReport> dataList,
                                                Func<DailyProductionReport, float?> getValue,
                                                Func<DailyProductionReport, float?> getWeight)
        {
            if (dataList == null || dataList.Count == 0) return 0;

            float weightedSum = 0;
            float totalWeight = 0;
            foreach (var d in dataList)
            {
                var value = getValue(d) ?? 0f;
                var weight = getWeight(d) ?? 0f;
                weightedSum += value * weight;
                totalWeight += weight;
            }
            return totalWeight == 0 ? 0 : weightedSum / totalWeight;
        }
        /***********************Excel需要计算的表***********************/
        /// <summary>
        /// 计算手写表一段时间的数据
        /// </summary>
        /// <param name="startTime">开始时间</param>
        /// <param name="endTime">结束时间</param>
        /// <returns>ProductionDataCollection</returns>
        public static List<DailyProductionReport> CalculateDailyProductionReportRange(DateTime startDate, DateTime endtDate, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            if (startDate > endtDate)
            {
                (startDate, endtDate) = (endtDate, startDate);
            }
            List<DailyProductionReport> DailyProductionReportCollection = [];
            var currentDay = startDate.Date.AddHours(8);
            var lastDay = endtDate.Date.AddHours(8);
            while (currentDay < lastDay)
            {
                var data = new DailyProductionReport(currentDay, sourceData, operatorInputData);
                if (data != null)
                {
                    DailyProductionReportCollection.AddRange(data);
                }
                currentDay = currentDay.AddDays(1);
            }

            return DailyProductionReportCollection;
        }

    }
}
