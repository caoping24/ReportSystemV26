namespace CenterBackend.Models.CalculateData
{
    //**********************通用方法**********************
    public static class MathTools
    {

        //**********************获取周第一天**********************
        public static DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = (int)dt.DayOfWeek - (int)DayOfWeek.Monday;
            if (diff < 0) diff += 7;
            return dt.AddDays(-diff).Date;
        }
        //**********************计算**********************
        //
        /// <summary>
        /// 计算 float? 类型平均值，自动忽略 null 值
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据集</param>
        /// <param name="selector">字段取值委托</param>
        /// <returns>平均值；无数据/无有效数据返回 null</returns>
        public static decimal? CalculateAverage<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            // 增加空集合判断
            if (data == null || !data.Any())
                return null;

            decimal sum = 0m;
            int count = 0;
            foreach (var item in data)
            {
                float? nullable = selector(item);
                if (nullable.HasValue)
                {
                    sum += (decimal)nullable.GetValueOrDefault();
                    count++;
                }
            }

            return count == 0 ? null : sum / count;
        }
        /// <summary>
        /// 计算 decimal? 类型平均值，自动忽略 null 值
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据集</param>
        /// <param name="selector">字段取值委托</param>
        /// <returns>平均值；无数据/无有效数据返回 null</returns>
        public static decimal? CalculateAverage<T>(IEnumerable<T> data, Func<T, decimal?> selector)
        {
            if (data == null || !data.Any())
                return null;

            decimal sum = 0m;
            int count = 0;
            foreach (var item in data)
            {
                decimal? nullable = selector(item);
                if (nullable.HasValue)
                {
                    sum += nullable.GetValueOrDefault();
                    count++;
                }
            }

            return count == 0 ? null : sum / count;
        }
        /// <summary>
        /// 计算 float? 类型集合首尾有效值的差值（最后一个 - 第一个），自动忽略 null 值
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据集</param>
        /// <param name="selector">字段取值委托</param>
        /// <returns>首尾差值；无数据/有效数据少于2个返回 null</returns>
        public static decimal? CalculateFirstLastDifference<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())
                return null;

            decimal firstValue = 0m;
            decimal lastValue = 0m;
            int validCount = 0;

            foreach (var item in data)
            {
                decimal? nullable = (decimal?)selector(item);
                if (nullable.HasValue)
                {
                    // 第一个有效值赋值给firstValue
                    if (validCount == 0)
                        firstValue = nullable.GetValueOrDefault();

                    // 每次都更新lastValue，最终保留最后一个有效值
                    lastValue = nullable.GetValueOrDefault();
                    validCount++;
                }
            }

            // 有效数据至少2个才返回差值
            return validCount >= 2 ? lastValue - firstValue : null;
        }

        /// <summary>
        /// 计算 decimal? 类型集合首尾有效值的差值（最后一个 - 第一个），自动忽略 null 值
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据集</param>
        /// <param name="selector">字段取值委托</param>
        /// <returns>首尾差值；无数据/有效数据少于2个返回 null</returns>
        public static decimal? CalculateFirstLastDifference<T>(IEnumerable<T> data, Func<T, decimal?> selector)
        {
            if (data == null || !data.Any())
                return null;

            decimal firstValue = 0m;
            decimal lastValue = 0m;
            int validCount = 0;

            foreach (var item in data)
            {
                decimal? nullable = selector(item);
                if (nullable.HasValue)
                {
                    if (validCount == 0)
                        firstValue = nullable.GetValueOrDefault();

                    lastValue = nullable.GetValueOrDefault();
                    validCount++;
                }
            }

            return validCount >= 2 ? lastValue - firstValue : null;
        }

        /// <summary>
        /// 计算 float? 类型集合的非空值总和，自动忽略 null 值，结果以 decimal? 类型返回（报表精度）
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据集</param>
        /// <param name="selector">float? 字段取值委托</param>
        /// <returns>非空值总和；无数据/无有效数据返回 null</returns>
        public static decimal? CalculateSum<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())
                return null;

            decimal sum = 0m;
            int validCount = 0;

            foreach (var item in data)
            {
                float? nullable = selector(item);
                if (nullable.HasValue)
                {
                    // 直接转decimal累加，避免float精度累积误差
                    sum += (decimal)nullable.GetValueOrDefault();
                    validCount++;
                }
            }

            return validCount > 0 ? sum : null;
        }
        /// <summary>
        /// 计算 decimal? 类型集合的非空值总和，自动忽略 null 值
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据集</param>
        /// <param name="selector">decimal? 字段取值委托</param>
        /// <returns>非空值总和；无数据/无有效数据返回 null</returns>
        public static decimal? CalculateSum<T>(IEnumerable<T> data, Func<T, decimal?> selector)
        {
            if (data == null || !data.Any())
                return null;

            decimal sum = 0m;
            int validCount = 0;

            foreach (var item in data)
            {
                decimal? nullable = selector(item);
                if (nullable.HasValue)
                {
                    sum += nullable.GetValueOrDefault();
                    validCount++;
                }
            }

            return validCount > 0 ? sum : null;
        }
        /// <summary>
        /// float 求反应时间
        /// </summary>
        /// <param name="d_mm"></param>
        /// <param name="L_m"></param>
        /// <param name="Q_L_per_h"></param>
        /// <returns></returns>
        public static decimal ResidenceTimeSeconds(decimal d_mm, decimal L_m, float Q_L_per_h)
        {
            if (Q_L_per_h <= 0f) return 0;

            decimal d_m = d_mm * 0.001m;
            decimal r_m = d_m / 2.0m;
            decimal area_m2 = (decimal)Math.PI * r_m * r_m;
            decimal V_m3 = area_m2 * (decimal)L_m;
            decimal V_L = V_m3 * 1000m;
            decimal t_h = V_L / (decimal)Q_L_per_h;
            decimal t_s = t_h * 3600m;

            return t_s;//保留两位小数
        }
        /// <summary>
        /// decimal 求反应时间
        /// </summary>
        /// <param name="d_mm"></param>
        /// <param name="L_m"></param>
        /// <param name="Q_L_per_h"></param>
        /// <returns></returns>
        public static decimal ResidenceTimeSeconds(decimal d_mm, decimal L_m, decimal Q_L_per_h)
        {
            if (Q_L_per_h <= 0m) return 0;

            decimal d_m = d_mm * 0.001m;
            decimal r_m = d_m / 2.0m;
            decimal area_m2 = (decimal)Math.PI * r_m * r_m;
            decimal V_m3 = area_m2 * (decimal)L_m;
            decimal V_L = V_m3 * 1000m;
            decimal t_h = V_L / Q_L_per_h;
            decimal t_s = t_h * 3600m;

            return t_s;//保留两位小数
        }
        //public static decimal CalculateWeighted(
        //                                        List<DailyProductionReport> dataList,
        //                                        Func<DailyProductionReport, float?> getValue,
        //                                        Func<DailyProductionReport, float?> getWeight)
        //{
        //    if (dataList == null || dataList.Count == 0) return 0;

        //    decimal weightedSum = 0;
        //    decimal totalWeight = 0;
        //    foreach (var d in dataList)
        //    {
        //        decimal value = (decimal)(getValue(d) ?? 0);
        //        decimal weight = (decimal)(getWeight(d) ?? 0);
        //        weightedSum += value * weight;
        //        totalWeight += weight;
        //    }
        //    return totalWeight == 0 ? 0 : weightedSum / totalWeight;
        //}




    }

}
