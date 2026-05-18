namespace CenterBackend.Models.CalculateData
{
    //**********************计算**********************
    public static class MathTools
    {
        //**********************通用方法**********************
        //求平均值，忽略 null 和非数值
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
        // 求两数差值，忽略 null 和非数值
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
        //// 列内求和
        //private static float SumRow(
        //                            List<ProductionData> dataList,
        //                            Func<ProductionData, float?> getValue)
        //{
        //    float Result = 0;
        //    if (dataList == null) return Result;
        //    foreach (var d in dataList)
        //    {
        //        float a = getValue(d) ?? 0f;
        //        Result += a;
        //    }
        //    return Result;
        //}
        //// 行内求和
        //private static void SumColumn(
        //                            List<ProductionData> dataList,
        //                            Func<ProductionData, float?> getValue1,
        //                            Func<ProductionData, float?> getValue2,
        //                            Action<ProductionData, float?> setValue3)
        //{
        //    if (dataList == null || dataList.Count == 0) return;
        //    // 逐行计算 a+b，收集有效结果
        //    List<float> rowResults = new List<float>();
        //    foreach (var d in dataList)
        //    {
        //        float a = getValue1(d) ?? 0f;
        //        float b = getValue2(d) ?? 0f;

        //        var Result = (a + b);
        //        rowResults.Add(Result);
        //        setValue3(d, Result);
        //    }
        //}
        //// 行内求折百
        //private static void MulColumn(
        //                        List<ProductionData> dataList,
        //                        Func<ProductionData, float?> getValue1,
        //                        Func<ProductionData, float?> getValue2,
        //                        Action<ProductionData, float?> setValue3)
        //{
        //    if (dataList == null || dataList.Count == 0) return;

        //    // 逐行计算 a*b/100，收集有效结果
        //    List<float> rowResults = new List<float>();
        //    foreach (var d in dataList)
        //    {
        //        float a = getValue1(d) ?? 0f;
        //        float b = getValue2(d) ?? 0f;

        //        var Result = (a * b) / 100;
        //        rowResults.Add(Result);
        //        setValue3(d, Result);
        //    }
        //}

        //// 所有列求两列加权平均
        //private static float WeightedAverageTowColumn(
        //                                            List<ProductionData> dataList,
        //                                            Func<ProductionData, float?> getValue,
        //                                            Func<ProductionData, float?> getWeight)
        //{
        //    if (dataList == null || dataList.Count == 0) return 0;

        //    float weightedSum = 0;
        //    float totalWeight = 0;
        //    foreach (var d in dataList)
        //    {
        //        var value = getValue(d) ?? 0f;
        //        var weight = getWeight(d) ?? 0f;
        //        weightedSum += value * weight;
        //        totalWeight += weight;
        //    }
        //    return totalWeight == 0 ? 0 : weightedSum / totalWeight;
        //}
        /// <summary>
        /// 反应时间计算公式 2026年5月13日增加
        /// </summary>
        /// <param name="d_mm">管内径</param>
        /// <param name="L_m">管长度</param>
        /// <param name="Q_L_per_h">流速L/H</param>
        /// <returns></returns>
        public static float ResidenceTimeSeconds(float d_mm, float L_m, float Q_L_per_h)
        {
            if (Q_L_per_h <= 0f) return 0f;

            double d_m = (double)d_mm * 1e-3;
            double r_m = d_m / 2.0;
            double area_m2 = Math.PI * r_m * r_m;
            double V_m3 = area_m2 * (double)L_m;
            double V_L = V_m3 * 1000.0;
            double t_h = V_L / (double)Q_L_per_h;
            double t_s = t_h * 3600.0;

            return (float)Math.Round(t_s, 3, MidpointRounding.AwayFromZero);//保留两位小数
        }



    }

}
