using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.Models;
using MathNet.Numerics.Optimization;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using static FastExpressionCompiler.ExpressionCompiler;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CenterBackend.Models.CalculateData
{
    //**********************数据结构**********************
    public class DailyProductionReport//sheet4
    {
        public int MaxBatches => 2; // 白班 + 夜班
        private readonly List<ShiftProductionData> _shiftBatches;
        public IReadOnlyList<ShiftProductionData> ShiftBatches => _shiftBatches;
        public DateTime TimePoint;
        public decimal? Cell1 { get; } // 当日总收率
        public decimal? Cell2 { get; } // 当日折百产量
        public decimal? Cell3 { get; } // 当日产量
        public DailyProductionReport(DateTime timeBase, List<SourceData> dataList1, List<OperatorInputData> dataList2)
        {
            _shiftBatches = new List<ShiftProductionData>();
            dataList1 ??= new List<SourceData>();
            dataList2 ??= new List<OperatorInputData>();
            // 白班 8:00 ~ 20:00
            var dayStart = timeBase.Date.AddHours(8);
            var dayEnd = dayStart.AddHours(13);
            var dayData1 = dataList1.Where(x => x.ReportedTime >= dayStart && x.ReportedTime < dayEnd).ToList();
            var dayData2 = dataList2.Where(x => x.ReportedTime >= dayStart && x.ReportedTime < dayEnd).ToList();
            _shiftBatches.Add(new ShiftProductionData(dayData1, dayData2));

            // 夜班 20:00 ~ 次日 8:00
            var nightStart = dayStart.AddHours(12);//当天20点
            var nightEnd = nightStart.AddHours(13);
            var nightData1 = dataList1.Where(x => x.ReportedTime >= nightStart && x.ReportedTime < nightEnd).ToList();
            var nightData2 = dataList2.Where(x => x.ReportedTime >= nightStart && x.ReportedTime < nightEnd).ToList();
            _shiftBatches.Add(new ShiftProductionData(nightData1, nightData2));

            // 构造时只计算一次（性能更好）
            Cell1 = CalculateYield();
            Cell2 = CalculateYieldWeight();
            Cell3 = CalculateTotalWeight();
            TimePoint = timeBase;
        }

        public decimal CalculateYield()
        {
            if (_shiftBatches.Count == 0) return 0;

            var batch1 = _shiftBatches[0];
            var batch2 = _shiftBatches[1];

            decimal a = batch1.Cell2;
            decimal b = batch1.Cell4;
            decimal c = batch2.Cell2;
            decimal d = batch2.Cell4;
            decimal e = batch1.Cell5;

            decimal denominator = b + d;
            if (denominator == 0 || e == 0) return 0;

            return (a + c) / denominator / e * 1.2m * 100m;
        }
        public decimal CalculateYieldWeight()
        {
            return _shiftBatches.Sum(b => b.Cell2);
        }
        public decimal CalculateTotalWeight()
        {
            return _shiftBatches.Sum(b => b.Cell1);
        }
    }
    public class ShiftProductionData
    {
        public int MaxBatches => 3;
        private readonly List<ProductionData> _batches = new();
        public IReadOnlyList<ProductionData> Batches => _batches;
        public decimal Cell1 => _batches.Sum(b => b.Cell3 ?? 0); // 产量累计
        public decimal Cell2 => _batches.Sum(b => b.Cell4 ?? 0); // 折百产量累计
        public decimal Cell3 => Cell2;
        public decimal Cell4 { get; } // 羟基用量
        public decimal Cell5 { get; } // 羟基浓度
        public decimal Cell6 => (Cell4 * Cell5 / 1000m); // 羟基折百
        public decimal Cell7 => Cell6;
        public decimal Cell8
        {
            get
            {
                if (Cell6 == 0) return 0;
                return Cell2 / Cell6 * 1.2m / 10m;
            }
        }

        public ShiftProductionData(List<SourceData> dataList1, List<OperatorInputData> dataList2)
        {
            Cell4 = MathTools.CalculateFirstLastDifference(dataList1, x => x.Cell20) /1000 ?? 0; //除以1000 L转立方
            var listWithoutLast = dataList1.Take(dataList1.Count - 1).ToList();
            Cell5 = MathTools.CalculateAverage(listWithoutLast, x => x.Cell6) ?? 0;

            if (dataList2 == null) return;

            foreach (var item in dataList2.Where(x => x != null))
            {
                var temp = new ProductionData(item);
                if (!temp.IsEmpty)
                    _batches.Add(temp);

                if (_batches.Count >= MaxBatches)
                    break;
            }
        }
    }
    public class ProductionData
    {
        public decimal? Cell1 { get; set; }
        public decimal? Cell2 { get; set; }
        public decimal? Cell3 { get; set; }
        public decimal? Cell4 => (Cell1.HasValue && Cell3.HasValue)
            ? Cell1.Value * Cell3.Value / 100
            : null;
        public ProductionData(OperatorInputData src)
        {
            if (src == null) return;

            Cell1 = (decimal?)src.Cell21;
            Cell2 = (decimal?)src.Cell23;
            Cell3 = (decimal?)src.Cell26;
        }
        public bool IsEmpty => Cell1 == null && Cell2 == null && Cell3 == null;
    }

}

