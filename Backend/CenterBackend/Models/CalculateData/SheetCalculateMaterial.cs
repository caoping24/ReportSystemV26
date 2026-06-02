using CenterReport.Repository.Models;

namespace CenterBackend.Models.CalculateData
{
    //**********************数据结构**********************
    //手写记录表 表6
    public class MaterialDailyCollection
    {
        public const int DataCount = 19;//需要计算的数据数量
        public List<MaterialData> MaterialDatas { get; private set; } = Enumerable.Range(0, DataCount).Select(_ => new MaterialData()).ToList();


        public decimal? Yield { get; private set; }//每天的折百产量
        public decimal? Usage { get; private set; }//每天的羟基消耗
        public decimal? Rate { get; private set; }//每天的羟基含量平均值
        public MaterialDailyCollection(DateTime startTime, decimal? yield, List<SourceData> sourceData, List<OperatorInputData> operatorInputData)
        {
            sourceData ??= new List<SourceData>();
            operatorInputData ??= new List<OperatorInputData>();

            Yield = yield;
            var start = startTime.Date.AddHours(8);
            var end = start.AddHours(25);//一天内范围 要多1个小时 才能包含25条数据

            var DataListFromDCS = sourceData.Where(x => x.ReportedTime >= start && x.ReportedTime < end).ToList();//左闭右开
            var DataListFromoperator = operatorInputData.Where(x => x.ReportedTime >= start && x.ReportedTime < end).ToList();//左闭右开
            foreach (var config in MaterialConfigs.AllItems)
            {
                decimal? value = null;

                // 根据配置的数据源 + 计算方式 自动计算
                if (config.DataSourceType == DataSourceType.DCS)
                {
                    value = config.CalculationType == CalculationType.FirstLastDifference
                        ? MathTools.CalculateFirstLastDifference(DataListFromDCS, config.DcsSelector!)
                        : MathTools.CalculateAverage(DataListFromDCS.Take(DataListFromDCS.Count - 1), config.DcsSelector!);//计算平均值剔除最后一个数据

                }
                else if (config.DataSourceType == DataSourceType.Operator)
                {
                    value = config.CalculationType == CalculationType.FirstLastDifference
                        ? MathTools.CalculateFirstLastDifference(DataListFromoperator, config.OperatorSelector!)
                        : MathTools.CalculateAverage(DataListFromoperator.Take(DataListFromoperator.Count - 1), config.OperatorSelector!);//计算平均值剔除最后一个数据

                }
                // 自动赋值到对应索引
                var eachItem = MaterialDatas[config.Index];
                eachItem.Index = config.Index;
                eachItem.CalculationType = config.CalculationType;
                eachItem.UsageOrAverage = value * config.Mul ?? 0; // 使用乘数修正单位
                eachItem.Specific = eachItem.UsageOrAverage;

                if (config.Index == 0) 
                    Usage = eachItem.UsageOrAverage;//获取羟基用量
                if (config.Index == 3)
                {
                    Rate = eachItem.UsageOrAverage;//获取羟基含量
                    MaterialDatas[0].Specific = MaterialDatas[0].Specific * Rate * 1000;//还要计算一下羟基的单耗,需要将Rate乘上(相当于单位是kg)
                }
            }
        }
}

public class MaterialData
    {
        public int Index { get; set; }
        public CalculationType CalculationType { get; set; }
        public decimal? UsageOrAverage { get; set; } = 0;
        public decimal? Specific { get; set; }

    }
    // 枚举定义（统一复用）
    public enum DataSourceType { DCS, Operator }
    public enum AggregationType { SumDiv, Average }// 周聚合类型（Sum/Average）
    public enum CalculationType { FirstLastDifference, Average }//数据收集类型  如果收集类型是  FirstLastDifference 每日统计要做单耗计算 如果是Average 则计算Average
    public class MaterialItemConfig
    {
        public int Index { get; set; }               // 索引
        public string Name { get; set; } = string.Empty; // 名称（便于维护）
        public AggregationType AggregationType { get; set; } // 周聚合类型（Sum/Average）
        public CalculationType CalculationType { get; set; } // 每日计算类型
        public DataSourceType DataSourceType { get; set; }   // 数据源类型
        public Func<SourceData, float?>? DcsSelector { get; set; } // DCS字段选择器
        public Func<OperatorInputData, float?>? OperatorSelector { get; set; } // 人工录入字段选择器
        public decimal Mul { get; set; } = 1;  // 单位修正乘数
    }
    public static class MaterialConfigs
    {
        public static readonly List<MaterialItemConfig> AllItems = new()
        {
            new MaterialItemConfig
            {
                Index = 0,
                Name = "羟基乙腈",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.FirstLastDifference,
                DcsSelector = x => x.Cell20, //用配后流量 可以方便和配后浓度一起计算单耗
                Mul = 0.001m,//L转立方

            },
            new MaterialItemConfig
            {
                Index = 1,
                Name = "液氨",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.FirstLastDifference,
                DcsSelector = x => x.Cell8,
                Mul = 1000,
            },
            new MaterialItemConfig
            {
                Index = 2,
                Name = "稀硫酸",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.FirstLastDifference,
                DcsSelector = x => x.Cell37,
                Mul = 1730,
            },
            new MaterialItemConfig
            {
                Index = 3,
                Name = "羟基浓度配料后",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell6,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 4,
                Name = "氨腈摩尔比",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell23,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 5,
                Name = "反应时间",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell20,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 6,
                Name = "反应压力",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell24,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 7,
                Name = "羟基加热温度",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell21,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 8,
                Name = "氨汽混合温度",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell17,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 9,
                Name = "管反热点温度",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell26,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 10,
                Name = "预冷器结晶温度",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell62,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 11,
                Name = "一次结晶温度",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell66,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 12,
                Name = "降膜蒸发温度",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell144,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 13,
                Name = "二次结晶温度",
                AggregationType = AggregationType.Average,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.Average,
                DcsSelector = x => x.Cell122,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 14,
                Name = "脱盐水消耗",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.FirstLastDifference,
                DcsSelector = x => x.Cell143,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 15,
                Name = "废液排放",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.DCS,
                CalculationType = CalculationType.FirstLastDifference,
                DcsSelector = x => x.Cell134,
                Mul = 1000,
            },
            new MaterialItemConfig
            {
                Index = 16,
                Name = "低压蒸汽",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.Operator,
                CalculationType = CalculationType.FirstLastDifference,
                OperatorSelector = x => x.Cell71,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 17,
                Name = "中压蒸汽",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.Operator,
                CalculationType = CalculationType.FirstLastDifference,
                OperatorSelector = x => x.Cell72,
                Mul = 1,
            },
            new MaterialItemConfig
            {
                Index = 18,
                Name = "电能消耗",
                AggregationType = AggregationType.SumDiv,
                DataSourceType = DataSourceType.Operator,
                CalculationType = CalculationType.FirstLastDifference,
                OperatorSelector = x => x.Cell73,
                Mul = 1,
            }
        };
    }
}
