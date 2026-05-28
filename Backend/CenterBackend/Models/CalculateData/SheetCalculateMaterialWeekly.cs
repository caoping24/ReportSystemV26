using CenterBackend.Models.CalculateData;
using CenterReport.Repository.Models;

namespace CenterBackend.Models.CalculateData
{

    // 继承 MaterialDataCollection， Specific 自动计算
    public class MaterialDataWeeklyCollection 
    {
        const int MaterialDataCount = 19;
        public IReadOnlyList<MaterialDailyCollection> DailyCollections { get; }
        public List<decimal> WeeklyCollections { get; private set; }
                     = Enumerable.Range(0, MaterialDataCount).Select(_ => default(decimal)).ToList();//每个条目一周的汇总
        public MaterialDataWeeklyCollection(DateTime monday,
                                            List<decimal?> yields,
                                            List<SourceData> sourceData,
                                            List<OperatorInputData> operatorInputData)
        {
            sourceData ??= new List<SourceData>();
            operatorInputData ??= new List<OperatorInputData>();
            yields ??= new List<decimal?>();

            // 1. 生成周一~周日共 7 个日对象
            var dailies = new List<MaterialDailyCollection>();
            for (int i = 0; i < 7; i++)
            {
                var dayYield = i < yields.Count ? yields[i] : null;
                dailies.Add(new MaterialDailyCollection(
                    monday.AddDays(i), dayYield, sourceData, operatorInputData));
            }
            DailyCollections = dailies.AsReadOnly();

            // 遍历配置，自动对7天数据做 求和/平均
            foreach (var config in MaterialConfigs.AllItems)
            {
                // 取出当前条目 7天的所有日数据
                var dayValues = dailies
                    .Select(d => d.MaterialDatas[config.Index].Specific)
                    .Where(x => x.HasValue)
                    .ToList();

                // 按配置自动聚合
                switch (config.AggregationType)
                {                    
                    case AggregationType.Sum:
                        WeeklyCollections[config.Index] = dayValues.Sum() ?? 0;
                        break;
                    case AggregationType.Average:
                        WeeklyCollections[config.Index] = dayValues.Count != 0 ? (dayValues.Where(x=>x!=0).Average() ?? 0) : 0; 
                        break;
                }
            }
        }
    }



}
