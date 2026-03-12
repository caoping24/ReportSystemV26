using CenterReport.Repository.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR.Protocol;
using NPOI.SS.Formula.Functions;
using static FastExpressionCompiler.ExpressionCompiler;

namespace CenterBackend.Models.SheetCalculateData
{

    public class ProductionDataCollection
    {
        public List<ProductionData> DayShiftData { get; set; } = [];
        public List<ProductionData> NightShiftData { get; set; } = [];
        public ProductionCalculationResult DayResult { get; set; } = new ProductionCalculationResult();
        public ProductionCalculationResult NightResult { get; set; } = new ProductionCalculationResult();
        public ProductionCalculationResult TotalResult { get; set; } = new ProductionCalculationResult();
    }
    public class ProductionData
    {
        public DateTime ReportedTime { get; set; }
        public float? Cell21 { get; set; }
        public float? Cell22 { get; set; }
        public float? Cell23 { get; set; }
        public float? Cell24 { get; set; }
        public float? Cell25 { get; set; }
        public float? Cell26 { get; set; }
        public float? Cell27 { get; set; }
        //
        public float? Cell31 { get; set; }
        public float? Cell32 { get; set; }
        public float? Cell33 { get; set; }
        public float? Cell34 { get; set; }
        public float? Cell35 { get; set; }
        public float? Cell36 { get; set; }
        public float? Cell37 { get; set; }
        //
        public float? Cell41 { get; set; }
        public float? Cell42 { get; set; }
        public float? Cell43 { get; set; }
        public float? Cell44 { get; set; }
        public float? Cell45 { get; set; }
        public float? Cell46 { get; set; }
        public float? Cell47 { get; set; }
        // 从OperatorInputData提取Cell21~Cell37
        public static ProductionData FromOperatorInput(OperatorInputData input)
        {
            if (input == null) return new ProductionData();
            return new ProductionData
            {
                ReportedTime = input.ReportedTime,

                Cell21 = input.Cell21,
                Cell22 = input.Cell22,
                Cell23 = input.Cell23,
                Cell24 = input.Cell24,
                Cell25 = input.Cell25,
                Cell26 = input.Cell26,
                Cell27 = input.Cell27,

                Cell31 = input.Cell31,
                Cell32 = input.Cell32,
                Cell33 = input.Cell33,
                Cell34 = input.Cell34,
                Cell35 = input.Cell35,
                Cell36 = input.Cell36,
                Cell37 = input.Cell37
            };
        }
    }

    public class ProductionCalculationResult
    {
        //一次
        public float FirstAverage_1 { get; set; }
        public float FirstAverage_2 { get; set; }
        public float FirstAverage_3 { get; set; }
        public float FirstAverage_4 { get; set; }
        public float FirstAverage_5 { get; set; }
        public float FirstProduction { get; set; }//总产量
        public float FirstYield { get; set; }//总折百产量
        //二次
        public float SecondAverage_1 { get; set; }
        public float SecondAverage_2 { get; set; }
        public float SecondAverage_3 { get; set; }
        public float SecondAverage_4 { get; set; }
        public float SecondAverage_5 { get; set; }
        public float SecondProduction { get; set; }//总产量
        public float SecondYield { get; set; }//总折百产量

        // 汇总结果
        public float AllAverage_1 { get; set; }
        public float AllAverage_2 { get; set; }
        public float AllAverage_3 { get; set; }
        public float AllAverage_4 { get; set; }
        public float AllAverage_5 { get; set; }
        public float AllProduction { get; set; }//总产量
        public float AllYield { get; set; }//总折百产量
    }

    //**********************计算**********************
    public static class ProductionDataCollectionExtensions
    {
        //整个表计算
        public static void CalculateSheet(this ProductionDataCollection collection)
        {
            collection.CalculateSingleCells();
            collection.CalculateTotalCells();
        }
        //**********************行内计算**********************
        private static void CalculateSingleCells(this ProductionDataCollection collection)
        {
            // 1. 白班
            //一次
            MulColumn(collection. DayShiftData, x => x.Cell21, x => x.Cell26, setValue3: (d, result) => d.Cell27 = result);
            //二次
            MulColumn(collection.DayShiftData, x => x.Cell31, x => x.Cell36, setValue3: (d, result) => d.Cell37 = result);
            //合计
            WeightedAverageFourColumn(collection.DayShiftData, x => x.Cell21, x => x.Cell26, x => x.Cell31, x => x.Cell36, setValue3: (d, result) => d.Cell41 = result);
            WeightedAverageFourColumn(collection.DayShiftData, x => x.Cell24, x => x.Cell26, x => x.Cell34, x => x.Cell36, setValue3: (d, result) => d.Cell44 = result);
            SumColumn(collection.DayShiftData, x => x.Cell26, x => x.Cell36, setValue3: (d, result) => d.Cell46 = result);
            SumColumn(collection.DayShiftData, x => x.Cell27, x => x.Cell37, setValue3: (d, result) => d.Cell47 = result);

            // 2. 夜班
            //一次
            MulColumn(collection.NightShiftData, x => x.Cell21, x => x.Cell26, setValue3: (d, result) => d.Cell27 = result);
            //二次
            MulColumn(collection.NightShiftData, x => x.Cell31, x => x.Cell36, setValue3: (d, result) => d.Cell37 = result);
            //合计
            WeightedAverageFourColumn(collection.NightShiftData, x => x.Cell21, x => x.Cell26, x => x.Cell31, x => x.Cell36, setValue3: (d, result) => d.Cell41 = result);
            WeightedAverageFourColumn(collection.NightShiftData, x => x.Cell24, x => x.Cell26, x => x.Cell34, x => x.Cell36, setValue3: (d, result) => d.Cell44 = result);
            SumColumn(collection.NightShiftData, x => x.Cell26, x => x.Cell36, setValue3: (d, result) => d.Cell46 = result);
            SumColumn(collection.NightShiftData, x => x.Cell27, x => x.Cell37, setValue3: (d, result) => d.Cell47 = result);
        }
        //**********************汇总计算**********************
        private static void CalculateTotalCells(this ProductionDataCollection collection)
        {
            //白班
            collection.DayResult.AllProduction = SumRow(collection.DayShiftData, x => x.Cell46);
            collection.DayResult.AllYield = SumRow(collection.DayShiftData, x => x.Cell47);
            collection.DayResult.AllAverage_1 = WeightedAverageTowColumn(collection.DayShiftData, x => x.Cell41, x => x.Cell46);
            collection.DayResult.AllAverage_4 = WeightedAverageTowColumn(collection.DayShiftData, x => x.Cell44, x => x.Cell46);

            //夜班
            collection.NightResult.AllProduction = SumRow(collection.NightShiftData, x => x.Cell46);
            collection.NightResult.AllYield = SumRow(collection.NightShiftData, x => x.Cell47);
            collection.NightResult.AllAverage_1 = WeightedAverageTowColumn(collection.NightShiftData, x => x.Cell41, x => x.Cell46);
            collection.NightResult.AllAverage_4 = WeightedAverageTowColumn(collection.NightShiftData, x => x.Cell44, x => x.Cell46);

            //当日
            var value1 = collection.DayResult.AllAverage_1 * collection.DayResult.AllProduction + collection.NightResult.AllAverage_1 * collection.NightResult.AllProduction;
            var value2 = collection.DayResult.AllAverage_4 * collection.DayResult.AllProduction + collection.NightResult.AllAverage_4 * collection.NightResult.AllProduction;
            var value3 = collection.DayResult.AllProduction +  collection.NightResult.AllProduction;
            var value4 = collection.DayResult.AllYield + collection.NightResult.AllYield;

            collection.TotalResult.AllAverage_1 = value1 / value3;
            collection.TotalResult.AllAverage_4 = value2 / value3;
            collection.TotalResult.AllProduction = value3;
            collection.TotalResult.AllYield = value4;
        }
        //**********************通用方法**********************
        // 列内求和
        private static float SumRow(
                                    List<ProductionData> dataList,
                                    Func<ProductionData, float?> getValue)
        {
            float Result = 0; 
            foreach (var d in dataList)
            {
                float a = getValue(d) ?? 0f;
                Result += a;
            }
            return Result;
        }
        // 行内求和
        private static void SumColumn(
                                    List<ProductionData> dataList,
                                    Func<ProductionData, float?> getValue1,
                                    Func<ProductionData, float?> getValue2,
                                    Action<ProductionData, float?> setValue3)
        {
            if (dataList == null || dataList.Count == 0) return;
            // 逐行计算 a+b，收集有效结果
            List<float> rowResults = [];
            foreach (var d in dataList)
            {
                float a = getValue1(d) ?? 0f;
                float b = getValue2(d) ?? 0f;

                var Result = (a + b);
                rowResults.Add(Result);
                setValue3(d, Result);
            }
        }
        // 行内求折百
        private static void MulColumn(
                                List<ProductionData> dataList,
                                Func<ProductionData, float?> getValue1,
                                Func<ProductionData, float?> getValue2,
                                Action<ProductionData, float?> setValue3)
        {
            if (dataList == null || dataList.Count == 0) return ;

            // 逐行计算 a*b/100，收集有效结果
            List<float> rowResults = [];
            foreach (var d in dataList)
            {
                float a = getValue1(d) ?? 0f;
                float b = getValue2(d) ?? 0f;

                var Result = (a * b) / 100;
                rowResults.Add(Result);
                setValue3(d, Result);
            }
        }



        // 所有列求两列加权平均
        private static float WeightedAverageTowColumn(
                                                    List<ProductionData> dataList, 
                                                    Func<ProductionData, float?> getValue, 
                                                    Func<ProductionData, float?> getWeight)
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
        // 行内求四列加权平均
        private static void WeightedAverageFourColumn(
                                                    List<ProductionData> dataList,
                                                    Func<ProductionData, float?> getValue1,
                                                    Func<ProductionData, float?> getWeight1,
                                                    Func<ProductionData, float?> getValue2,
                                                    Func<ProductionData, float?> getWeight2,
                                                    Action<ProductionData, float?> setValue3)
        {
            if (dataList == null || dataList.Count == 0) return;

            foreach (var d in dataList)
            {
                var value1 = getValue1(d) ?? 0f;
                var weight1 = getWeight1(d) ?? 0f;
                var value2 = getValue2(d) ?? 0f;
                var weight2 = getWeight2(d) ?? 0f;
                float weightedSum = (value1 * weight1) + (value2 * weight2);
                float totalWeight = (weight1 + weight2);
                setValue3(d, weightedSum / totalWeight);
            }
        }
        // 所有列四列加权平均
        private static float WeightedAverageFourColumn(
                                                    List<ProductionData> dataList,
                                                    Func<ProductionData, float?> getValue1,
                                                    Func<ProductionData, float?> getWeight1,
                                                    Func<ProductionData, float?> getValue2,
                                                    Func<ProductionData, float?> getWeight2)
        {
            if (dataList == null || dataList.Count == 0) return 0;

            float weightedSum = 0;
            float totalWeight = 0;
            foreach (var d in dataList)
            {
                var value1 = getValue1(d) ?? 0f;
                var weight1 = getWeight1(d) ?? 0f;
                var value2 = getValue2(d) ?? 0f;
                var weight2 = getWeight2(d) ?? 0f;
                weightedSum += (value1 * weight1)+ (value2 * weight2);
                totalWeight += (weight1 + weight2);
            }
            return totalWeight == 0 ? 0 : weightedSum / totalWeight;
        }
    }
}

