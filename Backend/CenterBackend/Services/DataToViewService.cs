using CenterBackend.IServices;
using CenterBackend.Models.CalculateData;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using Masuit.Tools;
using Masuit.Tools.Models;
using Microsoft.Identity.Client;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System.Collections;
using System.Security.Cryptography.Xml;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace CenterBackend.Services
{
    public class DataToViewService(IReportRepository<SourceData> sourceData,
                                    IReportRepository<OperatorInputData> operatorInputData) : IDataToViewService
    {

        private readonly IReportRepository<SourceData> _sourceData = sourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = operatorInputData;

        public async Task<bool> DayGetMapDataAsync(DayWorkBook DayWorkBook)
        {
            DateTime startTime = DayWorkBook.ReportedTime.Date.AddHours(8);
            DateTime endTime = startTime.AddDays(1).AddHours(8);

            var sourceData = await _sourceData.GetByDateTimeRangeAsync(startTime, endTime);
            if (sourceData == null || sourceData.Count == 0)
                return false;//未查到数据
            var operatorInputData = await _operatorInputData.GetByDateTimeRangeAsync(startTime, endTime);
            MoveDataSingleDay(DayWorkBook, sourceData, operatorInputData);
            MoveDataAnalysis(DayWorkBook, sourceData, operatorInputData);
            return true;
        }

        public async Task<bool> MonthGetMapDataAsync(MonthWorkBook monthWorkBook)
        {
            return false;
        }

        public async Task<bool> YearGetMapDataAsync(YearWorkBook yearWorkBook)
        {
            return false;
        }

        public async Task<bool> WeekGetMapDataAsync(WeekWorkBook WeekWorkBook)
        {
            DateTime currentYearFirstDay = new(WeekWorkBook.ReportedTime.Year, 1, 1);
            currentYearFirstDay= currentYearFirstDay.AddHours(8);
            DateTime currentYearLastDay = currentYearFirstDay.AddYears(1).AddDays(-1);
            List<SourceData> sourceData = await _sourceData.GetByDateTimeRangeAsync(currentYearFirstDay, currentYearLastDay);
            List< OperatorInputData > operatorInputData= await _operatorInputData.GetByDateTimeRangeAsync(currentYearFirstDay, currentYearLastDay);
            WeekMoveDataSheet2Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet3Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet4Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet5Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet6Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet7Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet8Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet9Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet10Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet11Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet12Async(WeekWorkBook, sourceData, operatorInputData);
            WeekMoveDataSheet13Async(WeekWorkBook, sourceData, operatorInputData);
            return true;
        }
        /***********************数据处理***********************/

        private static void MoveDataSingleDay(DayWorkBook DayWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            DayWorkBook.DaySheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();
            DayWorkBook.NightSheet = Enumerable.Range(0, 13).Select(_ => new DayWorkSheet()).ToList();

            var startTime = DayWorkBook.ReportedTime.Date.AddHours(8);
            var dataPart1 = SortDataByTime(sourceDatas, startTime, 25);//原始数据
            var dataPart2 = SortDataByTime(operatorInputDatas, startTime, 25);//人工录入数据

            List<SourceData> source1;
            List<OperatorInputData> source2;

            source1 = dataPart1.Take(13).ToList();
            source2 = dataPart2.Take(13).ToList();
            ShiftsMoveData(DayWorkBook.DaySheet, source1, source2);//白班

            source1 = dataPart1.Skip(12).Take(13).ToList();
            source2 = dataPart2.Skip(12).Take(13).ToList();
            ShiftsMoveData(DayWorkBook.NightSheet, source1, source2);//夜班
        }
        private static void MoveDataAnalysis(DayWorkBook DayWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            DateTime startTime = DayWorkBook.ReportedTime.Date.AddHours(8);
            DateTime shiftTime;
            ProductionDataCollection productionDataCollection = new();
            MaterialDataCollection materialDataCollection = new();

            DayWorkBook.AllDay = Enumerable.Range(0, 2).Select(_ => new DayAnalysis()).ToList();//白班+晚班
            List<DayAnalysis> target = DayWorkBook.AllDay;

            List<SourceData> sourcedataPart;
            List<OperatorInputData> operatorInputdataPart;
            if (sourceDatas == null || sourceDatas.Count == 0)
                return;
            if (operatorInputDatas == null || operatorInputDatas.Count == 0)
                return;
            //需要计算的数据
            {
                productionDataCollection = CalculateForSheet3(startTime, operatorInputDatas);
                float rangeYield = productionDataCollection.TotalResult.AllYield;//获取每日折百产量
                for (var y = 0; y < 10; y++)
                {
                    materialDataCollection.MaterialDatas[y].TotalResult.Yield = rangeYield;
                }
                materialDataCollection.CalculateSum();//计算原料单耗
            }
            //填充考评表数据
            for (var i = 0; i < 2; i++)
            {
                if (i != 0)
                {
                    startTime = startTime.AddHours(12);
                }
                shiftTime = startTime.AddHours(12);
                sourcedataPart = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < shiftTime).ToList();//区分白夜班
                operatorInputdataPart = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < shiftTime).ToList();//区分白夜班

                var shiftsProductionSum = productionDataCollection.DayResult.AllProduction + productionDataCollection.NightResult.AllProduction;
                var shiftsProductionAverage = shiftsProductionSum / 2;
                //第一列
                target[i].Cell1 = productionDataCollection.DayResult.AllProduction;
                target[i].Cell2 = productionDataCollection.DayResult.AllYield;
                //target[i].Cell3 = (productionDataCollection.DayResult.AllProduction >= shiftsProductionAverage) ? 5 : 6;//Wingdings 3 字体 5:上箭头 6 :下箭头   公式可以做在excel中
                target[i].Cell4 = productionDataCollection.DayResult.AllProduction - shiftsProductionAverage;
                target[i].Cell5 = productionDataCollection.DayResult.AllAverage_3;
                target[i].Cell6 = productionDataCollection.DayResult.AllAverage_5;

                target[i].Cell7 = productionDataCollection.DayResult.FirstProduction;
                target[i].Cell8 = productionDataCollection.DayResult.FirstYield;
                target[i].Cell9 = productionDataCollection.DayResult.FirstAverage_1;
                //target[i].Cell10 = ;
                //target[i].Cell11 = ;
                target[i].Cell12 = productionDataCollection.DayResult.FirstAverage_3;
                //target[i].Cell13 = ;
                //target[i].Cell14 = ;
                target[i].Cell15 = productionDataCollection.DayResult.FirstAverage_5;
                //target[i].Cell16 = ;
                //target[i].Cell17 = ;

                target[i].Cell18 = productionDataCollection.DayResult.SecondProduction;
                target[i].Cell19 = productionDataCollection.DayResult.SecondYield;
                target[i].Cell20 = productionDataCollection.DayResult.SecondAverage_1;
                target[i].Cell21 = productionDataCollection.DayResult.SecondAverage_3;
                target[i].Cell22 = productionDataCollection.DayResult.SecondAverage_5;
                target[i].Cell23 = (shiftsProductionSum != 0) ? productionDataCollection.DayResult.SecondProduction / shiftsProductionSum : 0; //占比

                if (materialDataCollection != null)
                {
                    target[i].Cell24 = (materialDataCollection.MaterialDatas[0] != null) ? materialDataCollection.MaterialDatas[0].TotalResult.Specific : 0f;
                    target[i].Cell25 = (materialDataCollection.MaterialDatas[1] != null) ? materialDataCollection.MaterialDatas[1].TotalResult.Specific : 0f;
                    target[i].Cell26 = (materialDataCollection.MaterialDatas[2] != null) ? materialDataCollection.MaterialDatas[2].TotalResult.Specific : 0f;
                    target[i].Cell27 = (materialDataCollection.MaterialDatas[3] != null) ? materialDataCollection.MaterialDatas[3].TotalResult.Specific : 0f;
                    target[i].Cell28 = (materialDataCollection.MaterialDatas[4] != null) ? materialDataCollection.MaterialDatas[4].TotalResult.Specific : 0f;
                    target[i].Cell29 = (materialDataCollection.MaterialDatas[5] != null) ? materialDataCollection.MaterialDatas[5].TotalResult.Specific : 0f;
                    target[i].Cell30 = (materialDataCollection.MaterialDatas[6] != null) ? materialDataCollection.MaterialDatas[6].TotalResult.Specific : 0f;
                    target[i].Cell31 = (materialDataCollection.MaterialDatas[7] != null) ? materialDataCollection.MaterialDatas[7].TotalResult.Specific : 0f;
                    target[i].Cell32 = (materialDataCollection.MaterialDatas[8] != null) ? materialDataCollection.MaterialDatas[8].TotalResult.Specific : 0f;
                    target[i].Cell33 = (materialDataCollection.MaterialDatas[9] != null) ? materialDataCollection.MaterialDatas[9].TotalResult.Specific : 0f;
  
                }
                //第二列
                //target[i].Cell34 = //手填基准值
                target[i].Cell35 = CalculateAverage(sourcedataPart, x => x.Cell24);//累计摩尔比
                //target[i].Cell36 = //公式做在模板中 累计摩尔比与手填基准值的差值  上下箭头表示

                target[i].Cell37 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell20) / 12;//累计除以当班时间
                target[i].Cell38 = CalculateAverage(sourcedataPart, x => x.Cell6);
                target[i].Cell39 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell14) / 12;
                target[i].Cell40 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell16) / 12;
                target[i].Cell41 = CalculateAverage(sourcedataPart, x => x.Cell25);
                target[i].Cell42 = CalculateAverage(sourcedataPart, x => x.Cell24);
                target[i].Cell43 = CalculateAverage(sourcedataPart, x => x.Cell26);
                target[i].Cell44 = CalculateAverage(sourcedataPart, x => x.Cell28);
                target[i].Cell45 = CalculateAverage(sourcedataPart, x => x.Cell27);

                //target[i].Cell46 = //固定值
                //target[i].Cell47 = //固定值
                target[i].Cell48 = CalculateAverage(operatorInputdataPart, x => x.Cell11);
                //target[i].Cell49 = 0;//excel中计算
                //target[i].Cell50 = 0;//excel中计算
                target[i].Cell51 = CalculateAverage(operatorInputdataPart, x => x.Cell18);
                //target[i].Cell52 = 0;//固定值
                //target[i].Cell53 = 0;//固定值
                target[i].Cell54 = CalculateAverage(operatorInputdataPart, x => x.Cell13);
                //target[i].Cell55 = 0;//excel中计算
                //target[i].Cell56 = 0;//excel中计算
                //target[i].Cell57 = 0;//固定值
                target[i].Cell58 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell37) / 12;
                //target[i].Cell59 = 0;//excel中计算
                //target[i].Cell60 = 0;//excel中计算
                target[i].Cell61 = CalculateAverage(sourcedataPart, x => x.Cell40);
                target[i].Cell62 = CalculateAverage(sourcedataPart, x => x.Cell39);
                target[i].Cell63 = CalculateAverage(sourcedataPart, x => x.Cell53);//闪发器液位没有统计  只有冷凝液罐的液位
                target[i].Cell64 = CalculateAverage(operatorInputdataPart, x => x.Cell1);
                target[i].Cell65 = CalculateAverage(operatorInputdataPart, x => x.Cell2);
                target[i].Cell66 = CalculateAverage(operatorInputdataPart, x => x.Cell3);
                //第三列
                target[i].Cell67 = CalculateAverage(operatorInputdataPart, x => x.Cell4);
                target[i].Cell68 = CalculateAverage(operatorInputdataPart, x => x.Cell5);
                
                target[i].Cell69 = CalculateAverage(sourcedataPart, x => x.Cell70);
                target[i].Cell70 = CalculateAverage(sourcedataPart, x => x.Cell66);
                //target[i].Cell71 = (target[i].Cell70 ?? 0f) - (target[i].Cell69 ?? 0f);//excel中计算差值
                target[i].Cell72 = CalculateAverage(sourcedataPart, x => x.Cell67);
                target[i].Cell73 = CalculateAverage(sourcedataPart, x => x.Cell65);
                target[i].Cell74 = productionDataCollection.DayResult.FirstYield;
                target[i].Cell75 = productionDataCollection.DayResult.FirstAverage_2;
                target[i].Cell76 = productionDataCollection.DayResult.FirstAverage_5;

                target[i].Cell77 = CalculateAverage(sourcedataPart, x => x.Cell125);
                target[i].Cell78 = CalculateAverage(sourcedataPart, x => x.Cell122);
                //target[i].Cell79 = (target[i].Cell78 ?? 0f) - (target[i].Cell77 ?? 0f);excel中计算差值
                target[i].Cell80 = CalculateAverage(sourcedataPart, x => x.Cell121);
                //target[i].Cell81 = //二次结晶没有PH
                target[i].Cell82 = productionDataCollection.DayResult.SecondYield;
                target[i].Cell83 = productionDataCollection.DayResult.SecondAverage_2;
                target[i].Cell84 = productionDataCollection.DayResult.SecondAverage_5;

                target[i].Cell85 = CalculateAverage(sourcedataPart, x => x.Cell110);
                target[i].Cell86 = CalculateAverage(sourcedataPart, x => x.Cell108);
                target[i].Cell87 = CalculateAverage(sourcedataPart, x => x.Cell92);
                //target[i].Cell88 =//低蒸液位没有统计
                target[i].Cell89 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell107);
                target[i].Cell90 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell107) / 12;
                target[i].Cell91 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell122);
                target[i].Cell92 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell122) / 12;
                target[i].Cell93 = CalculateAverage(sourcedataPart, x => x.Cell114);
                //target[i].Cell94 = 0;//浓缩比 excel中计算
                target[i].Cell95 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell112) / 12; ;
                target[i].Cell96 = CalculateAverage(sourcedataPart, x => x.Cell112);
                //float temp1, temp2, temp3;
                //temp1 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell107) ?? 0f;
                //temp2 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell116) ?? 0f;
                //temp3 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell112) ?? 0f;
                //target[i].Cell97 =  (temp1 - temp2) == 0 ? 0f : (temp3 / (temp1 - temp2));//低蒸蒸汽单耗 excel中计算
                //target[i].Cell98 = 0;//固定值
                //target[i].Cell99 = 0;//excel中计算


                target[i].Cell100 = CalculateAverage(operatorInputdataPart, x => x.Cell51);
                target[i].Cell101 = CalculateAverage(operatorInputdataPart, x => x.Cell57);
                //target[i].Cell102 = //excel中计算
                target[i].Cell103 = CalculateAverage(operatorInputdataPart, x => x.Cell53);
                target[i].Cell104 = CalculateAverage(operatorInputdataPart, x => x.Cell59);
                //target[i].Cell105 = //excel中计算
                target[i].Cell106 = CalculateAverage(operatorInputdataPart, x => x.Cell55);
                target[i].Cell107 = CalculateAverage(operatorInputdataPart, x => x.Cell61);
                //target[i].Cell108 = //excel中计算
                float? sum1, sum2;
                sum1 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell107) ?? 0f;
                sum2 = CalculateFirstLastDifference(sourcedataPart, x => x.Cell102) ?? 0f;
                target[i].Cell109 = (sum1 ?? 0f) - (sum2 ?? 0f);
                target[i].Cell110 = CalculateAverage(operatorInputdataPart, x => x.Cell64);
                //target[i].Cell111 = //excel中计算
                target[i].Cell112 = CalculateAverage(operatorInputdataPart, x => x.Cell62);
                //target[i].Cell113 = //固定值
            }

        }
        private static void ShiftsMoveData(List<DayWorkSheet> DayWorkSheet, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            var target = DayWorkSheet;
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
                        target[i].Cell112 = (currentVal - prevData) / 1000;
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
        private static bool WeekMoveDataSheet2Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet2 = Enumerable.Range(1, 3).Select(_ => new WorkSheet2()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = currentWeekFirstDay.AddDays(-14);
                    endTime = startTime.AddDays(7);
                }
                else if (i == 1)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }

                if (sourceDatas != null)
                { 
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();

                    WeekWorkBook.WorkSheet2[i].Cell1 = CalculateAverage(sourceData, x => x.Cell13);
                    WeekWorkBook.WorkSheet2[i].Cell2 = CalculateAverage(sourceData, x => x.Cell15);
                    WeekWorkBook.WorkSheet2[i].Cell3 = CalculateAverage(sourceData, x => x.Cell19);
                    WeekWorkBook.WorkSheet2[i].Cell4 = CalculateAverage(sourceData, x => x.Cell22);
                    WeekWorkBook.WorkSheet2[i].Cell5 = CalculateAverage(sourceData, x => x.Cell24);
                    WeekWorkBook.WorkSheet2[i].Cell6 = CalculateAverage(sourceData, x => x.Cell25);
                    WeekWorkBook.WorkSheet2[i].Cell7 = CalculateAverage(sourceData, x => x.Cell26);
                    WeekWorkBook.WorkSheet2[i].Cell8 = CalculateAverage(sourceData, x => x.Cell27);
                    WeekWorkBook.WorkSheet2[i].Cell9 = CalculateAverage(sourceData, x => x.Cell28);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet3Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet3 = Enumerable.Range(1, 3).Select(_ => new WorkSheet3()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else if (i == 1)
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Year, WeekWorkBook.ReportedTime.Month, 1).AddHours(8);//本月第一天
                    endTime = startTime.AddMonths(1).AddDays(-1);
                }
                if (operatorInputDatas != null)
                { 
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();

                    WeekWorkBook.WorkSheet3[i].Cell1 = CalculateAverage(operatorInputData, x => x.Cell11);
                    WeekWorkBook.WorkSheet3[i].Cell2 = CalculateAverage(operatorInputData, x => x.Cell12);
                    WeekWorkBook.WorkSheet3[i].Cell3 = CalculateAverage(operatorInputData, x => x.Cell13);
                    WeekWorkBook.WorkSheet3[i].Cell4 = CalculateAverage(operatorInputData, x => x.Cell14);
                    WeekWorkBook.WorkSheet3[i].Cell5 = CalculateAverage(operatorInputData, x => x.Cell17);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet4Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet4 = Enumerable.Range(1, 3).Select(_ => new WorkSheet4()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = currentWeekFirstDay.AddDays(-14);
                    endTime = startTime.AddDays(7);
                }
                else if (i == 1)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }

                if (sourceDatas != null)
                { 
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();

                    WeekWorkBook.WorkSheet4[i].Cell6 = CalculateAverage(sourceData, x => x.Cell92);
                    WeekWorkBook.WorkSheet4[i].Cell7 = CalculateAverage(sourceData, x => x.Cell106);
                }

                if (operatorInputDatas != null)
                { 
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();

                    WeekWorkBook.WorkSheet4[i].Cell1 = CalculateAverage(operatorInputData, x => x.Cell1);
                    WeekWorkBook.WorkSheet4[i].Cell2 = CalculateAverage(operatorInputData, x => x.Cell2);
                    WeekWorkBook.WorkSheet4[i].Cell3 = CalculateAverage(operatorInputData, x => x.Cell3);
                    WeekWorkBook.WorkSheet4[i].Cell4 = CalculateAverage(operatorInputData, x => x.Cell4);
                    WeekWorkBook.WorkSheet4[i].Cell5 = CalculateAverage(operatorInputData, x => x.Cell5);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet5Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet5 = Enumerable.Range(1, 3).Select(_ => new WorkSheet5()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Year, 1, 1).AddHours(8);//本年第一天
                    endTime = startTime.AddYears(1).AddDays(-1);

                }
                else if (i == 1)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }
                if (operatorInputDatas != null)
                { 
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
 
                    WeekWorkBook.WorkSheet5[i].Cell1 = CalculateAverage(operatorInputData, x => x.Cell21);
                    WeekWorkBook.WorkSheet5[i].Cell2 = CalculateAverage(operatorInputData, x => x.Cell23);
                    //WeekWorkBook.WorkSheet5[i].Cell3 = CalculateAverage(operatorInputData, x => x.Cell24);//自动计算
                    //WeekWorkBook.WorkSheet5[i].Cell4 = CalculateAverage(operatorInputData, x => x.Cell25);
                    WeekWorkBook.WorkSheet5[i].Cell5 = CalculateAverage(operatorInputData, x => x.Cell31);
                    WeekWorkBook.WorkSheet5[i].Cell6 = CalculateAverage(operatorInputData, x => x.Cell33);
                    //WeekWorkBook.WorkSheet5[i].Cell7 = CalculateAverage(operatorInputData, x => x.Cell34);//自动计算
                    //WeekWorkBook.WorkSheet5[i].Cell8 = CalculateAverage(operatorInputData, x => x.Cell35);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet6Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet6 = Enumerable.Range(1, 3).Select(_ => new WorkSheet6()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = currentWeekFirstDay.AddDays(-14);
                    endTime = startTime.AddDays(7);
                }
                else if (i == 1)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }
                var productionDataCollections = CalculateForSheet3TimeRange(startTime, endTime, operatorInputDatas);
                WeekWorkBook.WorkSheet6[i].Cell1 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_1);
                WeekWorkBook.WorkSheet6[i].Cell2 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_2);
            }
            return true;
        }
        private static bool WeekMoveDataSheet7Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet7 = Enumerable.Range(1, 14).Select(_ => new WorkSheet7()).ToList();
            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            for (var i = 0; i < 14; i++) 
            {
                if (i == 12)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else if (i == 13)
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = new DateTime(currentWeekFirstDay.Year, currentWeekFirstDay.Month, 1);
                    endTime = startTime.AddMonths(i).AddDays(-1);
                }
                var productionDataCollections =  CalculateForSheet3TimeRange(startTime, endTime, operatorInputDatas);
                WeekWorkBook.WorkSheet7[i].Cell1 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_1);
                WeekWorkBook.WorkSheet7[i].Cell2 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllProduction);
            }

            return true;
        }
        private static bool WeekMoveDataSheet8Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet8 = Enumerable.Range(1, 9).Select(_ => new WorkSheet8()).ToList();

            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            ProductionDataCollection ProductionDataCollection = new();
            for (var i = 0; i < 7; i++)
            {
                var startTime = currentWeekFirstDay.AddDays(i);
                var endTime = startTime.AddDays(1);
  

                if (sourceDatas != null)
                { 
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    WeekWorkBook.WorkSheet8[i].Cell1 = CalculateAverage(sourceData, x => x.Cell11);
                    WeekWorkBook.WorkSheet8[i].Cell2 = CalculateAverage(sourceData, x => x.Cell13);
                    WeekWorkBook.WorkSheet8[i].Cell3 = CalculateAverage(sourceData, x => x.Cell17);
                }
                if (operatorInputDatas != null)
                {
                    ProductionDataCollection = CalculateForSheet3(startTime, operatorInputDatas);
                    WeekWorkBook.WorkSheet8[i].Cell4 = ProductionDataCollection.TotalResult.AllAverage_1;
                    WeekWorkBook.WorkSheet8[i].Cell5 = ProductionDataCollection.TotalResult.AllAverage_3;
                    WeekWorkBook.WorkSheet8[i].Cell6 = ProductionDataCollection.TotalResult.AllAverage_4;

                    WeekWorkBook.WorkSheet8[i].Cell7 = ProductionDataCollection.TotalResult.AllProduction;
                    WeekWorkBook.WorkSheet8[i].Cell8 = ProductionDataCollection.TotalResult.AllYield;

                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    WeekWorkBook.WorkSheet8[i].Cell9 = CalculateAverage(operatorInputData, x => x.Cell63);
                }
            }
            for (var i = 7; i < 9; i++)
            {
                var startTime = currentWeekFirstDay;
                if (i == 7) startTime = currentWeekFirstDay.AddDays(-7);
                var endTime = startTime.AddDays(7);

                if (sourceDatas != null)
                {
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    WeekWorkBook.WorkSheet8[i].Cell1 = CalculateAverage(sourceData, x => x.Cell11);
                    WeekWorkBook.WorkSheet8[i].Cell2 = CalculateAverage(sourceData, x => x.Cell13);
                    WeekWorkBook.WorkSheet8[i].Cell3 = CalculateAverage(sourceData, x => x.Cell17);
                }
                List<ProductionDataCollection> productionDataCollections = [];
                if (operatorInputDatas != null)
                {
                    productionDataCollections = CalculateForSheet3TimeRange(startTime, endTime, operatorInputDatas);
                    WeekWorkBook.WorkSheet8[i].Cell4 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_1);
                    WeekWorkBook.WorkSheet8[i].Cell5 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_3);
                    WeekWorkBook.WorkSheet8[i].Cell6 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_4);
                    float temp = WeekWorkBook.WorkSheet8[i].Cell6 ?? 0f;
                    WeekWorkBook.WorkSheet8[i].Cell7 = 100 - temp;
                    WeekWorkBook.WorkSheet8[i].Cell8 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllProduction);
                    WeekWorkBook.WorkSheet8[i].Cell9 = CalculateAverage(productionDataCollections, x => x.TotalResult.AllYield);

                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    WeekWorkBook.WorkSheet8[i].Cell10 = CalculateAverage(operatorInputData, x => x.Cell63);
                }

            }

            return true;
        }
        private static bool WeekMoveDataSheet9Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet9 = Enumerable.Range(1, 2).Select(_ => new WorkSheet9()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 2; i++)
            {
                if (i == 0)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }
                if (sourceDatas != null)
                {
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    WeekWorkBook.WorkSheet9[i].Cell1 = CalculateFirstLastDifference(sourceData, x => x.Cell107);
                    WeekWorkBook.WorkSheet9[i].Cell2 = CalculateAverage(sourceData, x => x.Cell116);
                    float input = WeekWorkBook.WorkSheet9[i].Cell1 ?? 0f;
                    float output = WeekWorkBook.WorkSheet9[i].Cell2 ?? 0f;
                    if (output != 0) WeekWorkBook.WorkSheet9[i].Cell3 = input / output;

                    WeekWorkBook.WorkSheet9[i].Cell4 = CalculateAverage(sourceData, x => x.Cell114);
                    WeekWorkBook.WorkSheet9[i].Cell5 = CalculateAverage(sourceData, x => x.Cell112);
                }
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    WeekWorkBook.WorkSheet9[i].Cell6 = CalculateAverage(operatorInputData, x => x.Cell41);
                    WeekWorkBook.WorkSheet9[i].Cell7 = CalculateAverage(operatorInputData, x => x.Cell43);
                    WeekWorkBook.WorkSheet9[i].Cell8 = CalculateAverage(operatorInputData, x => x.Cell45);
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet10Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet10 = Enumerable.Range(1, 2).Select(_ => new WorkSheet10()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 2; i++)
            {
                if (i == 0)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }
                float? total = 0;
                if (sourceDatas != null)
                {
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (sourceData != null && sourceData.Count != 0)
                    {
                        total = CalculateSum(sourceData, x => x.Cell105);//活性炭消耗总量
                    }
                }
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    if (total != null)
                    {
                        var difference = CalculateFirstLastDifference(operatorInputData, x => x.Cell64);
                        if (difference != null)
                            WeekWorkBook.WorkSheet10[i].Cell1 = difference == 0 ? 0f : (total / difference);//活性炭单耗
                    }
                    WeekWorkBook.WorkSheet10[i].Cell2 = CalculateAverage(operatorInputData, x => x.Cell62);
                    if (i != 0)
                    {
                        WeekWorkBook.WorkSheet10[i].Cell3 = CalculateAverage(operatorInputData, x => x.Cell51);
                        WeekWorkBook.WorkSheet10[i].Cell4 = CalculateAverage(operatorInputData, x => x.Cell52);
                        WeekWorkBook.WorkSheet10[i].Cell5 = CalculateAverage(operatorInputData, x => x.Cell55);
                        WeekWorkBook.WorkSheet10[i].Cell6 = CalculateAverage(operatorInputData, x => x.Cell57);
                        WeekWorkBook.WorkSheet10[i].Cell7 = CalculateAverage(operatorInputData, x => x.Cell59);
                        WeekWorkBook.WorkSheet10[i].Cell8 = CalculateAverage(operatorInputData, x => x.Cell61);
                    }
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet11Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet11 = Enumerable.Range(1, 3).Select(_ => new WorkSheet11()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            for (var i = 0; i < 3; i++)
            {
                if (i == 0)
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Date.Year, WeekWorkBook.ReportedTime.Date.Month, 1).AddHours(8);
                    endTime = startTime.AddMonths(1).AddDays(-1);
                }
                else if (i == 1)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }

                if (sourceDatas != null)
                {
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    WeekWorkBook.WorkSheet11[i].Cell4 = CalculateFirstLastDifference(sourceData, x => x.Cell132);//废液外排累计
                }

                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    var productionDataCollections = CalculateForSheet3TimeRange(startTime, endTime, operatorInputData);

                    WeekWorkBook.WorkSheet11[i].Cell1 = CalculateAverage(operatorInputData, x => x.Cell52);
                    WeekWorkBook.WorkSheet11[i].Cell2 = CalculateAverage(operatorInputData, x => x.Cell53);
                    WeekWorkBook.WorkSheet11[i].Cell3 = CalculateAverage(operatorInputData, x => x.Cell55);

                    var sum = CalculateSum(productionDataCollections, x => x.TotalResult.AllProduction);
                    WeekWorkBook.WorkSheet11[i].Cell4 = sum == 0 ? 0 : WeekWorkBook.WorkSheet11[i].Cell4 / sum;//废液外排单耗
                    WeekWorkBook.WorkSheet11[i].Cell5 = sum;//废液外排累计
                }
            }
            return true;
        }
        private static bool WeekMoveDataSheet12Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            WeekWorkBook.WorkSheet12 = Enumerable.Range(1, 3).Select(_ => new WorkSheet12()).ToList();

            DateTime startTime;
            DateTime endTime;
            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);

            ProductionDataCollection productionDataCollection = new();
            MaterialDataCollection materialDataCollection = new();
            for (var i = 0; i < 3; i++) 
            {
                if (i == 0)
                {
                    startTime = new DateTime(WeekWorkBook.ReportedTime.Date.Year, WeekWorkBook.ReportedTime.Date.Month, 1).AddHours(8);
                    endTime = startTime.AddMonths(1).AddDays(-1);
                }
                else if (i == 1)
                {
                    startTime = currentWeekFirstDay.AddDays(-7);
                    endTime = startTime.AddDays(7);
                }
                else
                {
                    startTime = currentWeekFirstDay;
                    endTime = startTime.AddDays(7);
                }

                float rangeYield = 0;
                if (operatorInputDatas != null)
                {
                    productionDataCollection = CalculateForSheet3(startTime, operatorInputDatas);
                    rangeYield = productionDataCollection.TotalResult.AllYield;//获取每日折百产量
                }
                for (var y = 0; y < 10; y++)
                {
                    materialDataCollection.MaterialDatas[y].TotalResult.Yield = rangeYield;
                }

                if (sourceDatas != null)
                {
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    var temp = CalculateFirstLastDifference(sourceData, x => x.Cell4);
                    materialDataCollection.MaterialDatas[0].TotalResult.Usage = temp;
                    materialDataCollection.MaterialDatas[1].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell8);
                    materialDataCollection.MaterialDatas[2].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell37);
                    //materialDataCollection.MaterialDatas[3].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell4);
                    materialDataCollection.MaterialDatas[4].TotalResult.Usage = temp * 0.18218f / 1000;
                    materialDataCollection.MaterialDatas[5].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell112);
                    //vmaterialDataCollection.MaterialDatas[6].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell4);
                    //materialDataCollection.MaterialDatas[7].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell4);
                    materialDataCollection.MaterialDatas[8].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell55) + CalculateFirstLastDifference(sourceData, x => x.Cell118);
                    materialDataCollection.MaterialDatas[9].TotalResult.Usage = CalculateFirstLastDifference(sourceData, x => x.Cell134);
                }
                if (operatorInputDatas != null)
                {
                    var operatorInputData = operatorInputDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
                    materialDataCollection.MaterialDatas[3].TotalResult.Usage = CalculateFirstLastDifference(operatorInputData, x => x.Cell71);
                    materialDataCollection.MaterialDatas[6].TotalResult.Usage = CalculateFirstLastDifference(operatorInputData, x => x.Cell72);
                    materialDataCollection.MaterialDatas[7].TotalResult.Usage = CalculateFirstLastDifference(operatorInputData, x => x.Cell73);
                }
                materialDataCollection.CalculateSum();

                for (var j = 0; j < WeekWorkBook.WorkSheet12.Count; j++)
                {
                    WeekWorkBook.WorkSheet12[j].Cell1 = materialDataCollection.MaterialDatas[0].TotalResult.Yield;
                    WeekWorkBook.WorkSheet12[j].Cell2 = materialDataCollection.MaterialDatas[1].TotalResult.Yield;
                    WeekWorkBook.WorkSheet12[j].Cell3 = materialDataCollection.MaterialDatas[2].TotalResult.Yield;
                    WeekWorkBook.WorkSheet12[j].Cell4 = materialDataCollection.MaterialDatas[3].TotalResult.Yield;
                    WeekWorkBook.WorkSheet12[j].Cell5 = materialDataCollection.MaterialDatas[4].TotalResult.Yield;
                    WeekWorkBook.WorkSheet12[j].Cell6 = materialDataCollection.MaterialDatas[5].TotalResult.Yield;
                    WeekWorkBook.WorkSheet12[j].Cell7 = materialDataCollection.MaterialDatas[8].TotalResult.Yield;
                    WeekWorkBook.WorkSheet12[j].Cell8 = materialDataCollection.MaterialDatas[9].TotalResult.Yield;
                }
            }

            return true;
        }
        private static bool WeekMoveDataSheet13Async(WeekWorkBook WeekWorkBook, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {

            WeekWorkBook.WorkSheet13 = Enumerable.Range(1, 14).Select(_ => new WorkSheet13()).ToList();

            DateTime currentWeekFirstDay = GetWeekFirstDay(WeekWorkBook.ReportedTime.Date).AddHours(8);
            ProductionDataCollection ProductionDataCollection= new();
            for (var i = 0; i < 7; i++)
            {
                var startTime = currentWeekFirstDay.AddDays(i);
                var endTime = startTime.AddDays(1);

                if (operatorInputDatas != null)
                {
                    ProductionDataCollection =  CalculateForSheet3(startTime, operatorInputDatas);
                }

                if (sourceDatas != null)
                {
                    var sourceData = sourceDatas.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();

                    var dayShift = sourceData.Where(x => x.ReportedTime < startTime.AddHours(12));
                    WeekWorkBook.WorkSheet13[2 * i].TimePoint = startTime;
                    WeekWorkBook.WorkSheet13[2 * i].Cell1 = CalculateQualifiedRate(dayShift, x => x.Cell23, true, 0.515f, 0.05f);
                    WeekWorkBook.WorkSheet13[2 * i].Cell2 = CalculateQualifiedRate(dayShift, x => x.Cell3, true, 410f, 5f);
                    WeekWorkBook.WorkSheet13[2 * i].Cell3 = CalculateQualifiedRate(dayShift, x => x.Cell6, true, 168f, 2f);
                    WeekWorkBook.WorkSheet13[2 * i].Cell4 = CalculateQualifiedRate(dayShift, x => x.Cell66, false, 20, 0);
                    //WeekWorkBook.WorkSheet13[2 * i].Cell5 = 
                    WeekWorkBook.WorkSheet13[2 * i].Cell6 = ProductionDataCollection.DayResult.AllProduction;
                    WeekWorkBook.WorkSheet13[2 * i].Cell7 = ProductionDataCollection.DayResult.AllYield;
                    WeekWorkBook.WorkSheet13[2 * i].Cell8 = ProductionDataCollection.DayResult.AllAverage_1;

                    var nightShift = sourceData.Where(x => x.ReportedTime >= startTime.AddHours(12));
                    WeekWorkBook.WorkSheet13[2 * i + 1].TimePoint = startTime.AddHours(12);
                    WeekWorkBook.WorkSheet13[2 * i + 1].Cell1 = CalculateQualifiedRate(nightShift, x => x.Cell23, true, 0.515f, 0.05f);
                    WeekWorkBook.WorkSheet13[2 * i + 1].Cell2 = CalculateQualifiedRate(nightShift, x => x.Cell3, true, 410f, 5f);
                    WeekWorkBook.WorkSheet13[2 * i + 1].Cell3 = CalculateQualifiedRate(nightShift, x => x.Cell6, true, 168f, 2f);
                    WeekWorkBook.WorkSheet13[2 * i + 1].Cell4 = CalculateQualifiedRate(nightShift, x => x.Cell66, false, 20, 0);
                    //WeekWorkBook.WorkSheet13[2 * i + 1].Cell5 =
                    WeekWorkBook.WorkSheet13[2 * i + 1].Cell6 = ProductionDataCollection.NightResult.AllProduction;
                    WeekWorkBook.WorkSheet13[2 * i + 1].Cell7 = ProductionDataCollection.NightResult.AllYield;
                    WeekWorkBook.WorkSheet13[2 * i + 1].Cell8 = ProductionDataCollection.NightResult.AllAverage_1;
                }
            }
            return true;
        }
        /***********************辅助方法***********************/
        private static List<SourceData> SortDataByTime(List<SourceData> sourceData, DateTime baseDate, int maxCount)
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
        private static DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = (int)dt.DayOfWeek - (int)DayOfWeek.Monday;
            if (diff < 0) diff += 7;
            return dt.AddDays(-diff).Date;
        }
        private static float? CalculateAverage<T>(IEnumerable<T> data, Func<T, float?> selector)
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
        private static float? CalculateFirstLastDifference<T>(IEnumerable<T> data, Func<T, float?> selector)
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
        /// <param name="mode">mode：模式开关（true = 区间校验 [min,max]，false = 上限校验 [≤max]）</param>
        private static float? CalculateQualifiedRate<T>(IEnumerable<T> data, Func<T, float?> selector, bool mode, float qualifiedValue, float qualifiedValuediff)
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
            float qualifiedRate = (qualifiedCount / (float)nonNullValues.Count) * 100f;// 计算合格利率（合格数/总有效数 * 100%），保留3位小数
            return (float)Math.Round(qualifiedRate, 3);
        }
        private static float? CalculateSum<T>(IEnumerable<T> data, Func<T, float?> selector)//非null值的总和
        {
            if (data == null || !data.Any())//空数据校验
                return null;
            var nonNullValues = data        //筛选非null的float值
                .Select(selector)           // 提取float?字段
                .Where(x => x.HasValue)     // 过滤掉null值
                .Select(x => x.GetValueOrDefault())      // 转换为float（非可空）
                .ToList();
            return nonNullValues.Count != 0 ? nonNullValues.Sum() : (float?)null;//计算总和
        }
        /***********************Excel***********************/

        /// <summary>
        /// 计算手写表一天的数据
        /// </summary>
        /// <param name="startTime">对应当天日期</param>
        /// <returns>返回计算完成的sheet3的数据集合</returns>
        private static ProductionDataCollection CalculateForSheet3(DateTime startTime, List<OperatorInputData> operatorInputData)
        {
            //查询当日数据
            startTime = startTime.Date.AddHours(8);
            var endTime = startTime.AddDays(24);

            operatorInputData = operatorInputData.Where(x => x.ReportedTime >= startTime && x.ReportedTime < endTime).ToList();
            if (operatorInputData == null || operatorInputData.Count == 0)// 空数据校验
                return new ProductionDataCollection();

            //填充当日数据
            ProductionDataCollection dataCellection = new();
            IEnumerable<OperatorInputData> filteredList;
            var ShiftStart = startTime;//早班开始
            var ShiftTime = ShiftStart.AddHours(12);//换班时间
            var ShiftEnd = ShiftTime.AddHours(12);//换班时间

            filteredList = operatorInputData.Where(x => x.ReportedTime >= ShiftStart && x.ReportedTime < ShiftTime && x.Cell21 != null).Take(5);
            var DayShiftData = (filteredList != null) ? filteredList.ToList() : [];//早班数据
            dataCellection.DayShiftData = DayShiftData.Select(ProductionData.FromOperatorInput).ToList();

            filteredList = operatorInputData.Where(x => x.ReportedTime >= ShiftTime && x.ReportedTime < ShiftEnd && x.Cell21 != null).Take(5);
            var NightShiftData = (filteredList != null) ? filteredList.ToList() : [];//晚班数据
            dataCellection.NightShiftData = NightShiftData.Select(ProductionData.FromOperatorInput).ToList();
            
            if (dataCellection.DayShiftData == null || dataCellection.DayShiftData.Count == 0) return dataCellection;
            if (dataCellection.NightShiftData == null || dataCellection.NightShiftData.Count == 0) return dataCellection;
            //表内计算,得出当日所有统计数据
            dataCellection.CalculateSheet();
            return dataCellection;
        }

        /// <summary>
        /// 计算手写表一段时间的数据
        /// </summary>
        /// <param name="startTime">开始时间</param>
        /// <param name="endTime">结束时间</param>
        /// <returns>ProductionDataCollection</returns>
        private static List<ProductionDataCollection> CalculateForSheet3TimeRange(DateTime startDate, DateTime endtDate, List<OperatorInputData> operatorInputData)
        {
            if (startDate > endtDate)
            {
                (startDate, endtDate) = (endtDate, startDate);
            }
            List<ProductionDataCollection> productionDataCollection = [];
            var currentDay = startDate.Date.AddHours(8);
            var lastDay = endtDate.Date.AddHours(8);
            while (currentDay <= lastDay)
            {
                var data = CalculateForSheet3(currentDay, operatorInputData);
                if (data != null)
                {
                    if (data.DayShiftData.Count != 0) productionDataCollection.AddRange(data);
                }
                currentDay = currentDay.AddDays(1);
            }

            return productionDataCollection;
        }


    }
}
