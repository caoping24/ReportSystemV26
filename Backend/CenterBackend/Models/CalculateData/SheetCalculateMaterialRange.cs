using CenterBackend.Models.CalculateData;
using CenterReport.Repository.Models;

namespace CenterBackend.Models.CalculateData
{

    // 继承 MaterialDataCollection， Specific 自动计算
    public class MaterialDataRangeCollection
    {
        // 单条物料数据总条数（固定）
        private const int MaterialDataCount = 19;

        public IReadOnlyList<MaterialDailyCollection> DailyCollections { get; }
        // 周期汇总集合
        public List<decimal> RangeCollections { get; private set; }
            = Enumerable.Range(0, MaterialDataCount).Select(_ => default(decimal)).ToList();

        /// <summary>
        /// 构造函数：支持**任意天数**的周期数据汇总
        /// </summary>
        /// <param name="startDate">起始日期</param>
        /// <param name="totalDays">总天数（至少1天）</param>
        /// <param name="yields">每日收率数据</param>
        /// <param name="sourceData">原始数据源</param>
        /// <param name="operatorInputData">操作录入数据</param>
        public MaterialDataRangeCollection(
            DateTime startDate,
            int totalDays,
            List<decimal?> yields,
            List<SourceData> sourceData,
            List<OperatorInputData> operatorInputData)
        {
            //天数最小为1
            int actualDays = Math.Max(1, totalDays);

            // 空列表容错
            sourceData ??= new List<SourceData>();
            operatorInputData ??= new List<OperatorInputData>();
            yields ??= new List<decimal?>();

            var dailies = new List<MaterialDailyCollection>();
            for (int i = 0; i < actualDays; i++)
            {
                // 超出yields长度则赋值null
                var dayYield = i < yields.Count ? yields[i] : null;
                var currentDate = startDate.AddDays(i);
                dailies.Add(new MaterialDailyCollection(currentDate, dayYield, sourceData, operatorInputData));
            }
            DailyCollections = dailies.AsReadOnly();

            foreach (var config in MaterialConfigs.AllItems)
            {
                var dayValues = dailies
                    .Select(d => d.MaterialDatas[config.Index].Specific)
                    .Where(x => x.HasValue)
                    .ToList();

                // 按配置自动聚合
                switch (config.AggregationType)
                {
                    case AggregationType.Sum:
                        RangeCollections[config.Index] = dayValues.Sum() ?? 0;
                        break;
                    case AggregationType.Average:
                        RangeCollections[config.Index] = dayValues.Count != 0 ? (dayValues.Where(x => x != 0).Average() ?? 0) : 0;
                        break;
                }
            }
        }
    }
}
