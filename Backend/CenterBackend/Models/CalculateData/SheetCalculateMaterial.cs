using System.Collections;

namespace CenterBackend.Models.CalculateData
{
    //**********************数据结构**********************
    //手写记录表 表6
    public class MaterialDataCollection 
    {
        public List<SingleDay> MaterialDatas { get; set; } = Enumerable.Range(0, 10).Select(_ => new SingleDay()).ToList();
    }

    public class SingleDay
    {
        public MaterialData DayShift { get; set; } = new MaterialData();
        public MaterialData NightShift { get; set; } = new MaterialData();
        public MaterialData TotalResult { get; set; } = new MaterialData();

    }
    public class MaterialData
    {
        private float? _usage = 0;
        private float? _yield = 0;
        private float? _specific = null;
        public float? Usage
        {
            get => _usage;
            set
            {
                if (_usage != value)
                {
                    _usage = value;
                    CalculateSpecific();
                }
            }
        }
        public float? Yield
        {
            get => _yield;
            set
            {
                if (_yield != value)// 自动触发计算
                {
                    _yield = value;
                    CalculateSpecific();
                }
            }
        }
        public float? Specific//内部计算赋值
        {
            get => _specific;
            private set => _specific = value; // 禁止外部直接修改
        }
        private void CalculateSpecific()
        {
            if (Usage.HasValue && Yield.HasValue && Yield.Value != 0)
            {
                Specific = Usage.Value / Yield.Value;
            }
            else
            {
                Specific = null;
            }
        }
    }

    public static class MaterialDataCollectionExtensions
    {
        public static void CalculateSum(this MaterialDataCollection collection)
        {
            if (collection == null)
                throw new ArgumentNullException(nameof(collection), "物料数据集合不能为空");

            if (collection.MaterialDatas == null)// 空列表防护：避免遍历null列表
                return;
            foreach (var singleDay in collection.MaterialDatas)
            {
                singleDay.TotalResult.Usage =
                    (singleDay.DayShift.Usage ?? 0) + (singleDay.NightShift.Usage ?? 0);// TotalResult.Usage = 白班Usage + 夜班Usage

                singleDay.TotalResult.Yield =
                    (singleDay.DayShift.Yield ?? 0) + (singleDay.NightShift.Yield ?? 0);// TotalResult.Yield = 白班Yield + 夜班Yield
            }
        }
    }
}
