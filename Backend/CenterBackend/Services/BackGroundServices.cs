using CenterBackend.IServices;
using CenterBackend.Models.CalculateData;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;


namespace CenterBackend.Services
{


    public class BackGroundServices(IReportRepository<SourceData> sourceData,
                                    IReportRepository<OperatorInputData> operatorInputData,
                                    IReportRepository<ComToSiemens> comToSiemens,
                                    IReportUnitOfWork reportUnitOfWork) : IBackGroundServices
    {
        private readonly IReportRepository<SourceData> _sourceData = sourceData;
        private readonly IReportRepository<OperatorInputData> _operatorInputData = operatorInputData;
        private readonly IReportRepository<ComToSiemens> _comToSiemens = comToSiemens;
        private readonly IReportUnitOfWork _reportUnitOfWork = reportUnitOfWork;

        public async Task Daily0810()
        {
            DateTime start = DateTime.Now.Date.AddHours(8);
            DateTime end = start.AddHours(25); //25条数据
            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(start, end);
            List<OperatorInputData> operatorInputDatas = await _operatorInputData.GetByDateTimeRangeAsync(start, end);
            var data = await CalDataAsync(1, start, end, sourceDatas, operatorInputDatas);
            if (data != null)
            {
                await _comToSiemens.AddAsync(data);
                await _reportUnitOfWork.SaveChangesAsync(); ;
            }
        }
        public async Task WeeklyMon0820()
        {
            DateTime today = DateTime.Now.Date; // 缓存当前日期，避免跨零点时出现不一致
            DateTime start = today.AddDays(-7).AddHours(8); // 7天前的08:00:00
            DateTime end = today.AddHours(9); // 今天的09:00:00
            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(start, end);
            List<OperatorInputData> operatorInputDatas = await _operatorInputData.GetByDateTimeRangeAsync(start, end);
            var data = await CalDataAsync(2, start, end, sourceDatas, operatorInputDatas);
            if (data != null)
            {
                await _comToSiemens.AddAsync(data);
                await _reportUnitOfWork.SaveChangesAsync(); ;
            }
        }
        public async Task MonthlyDay1_0830()
        {
            DateTime today = DateTime.Now.Date;
            DateTime start = today.AddDays(1 - today.Day) // 本月1号
                              .AddMonths(-1) // 上个月1号
                              .AddHours(8); // 8点整
            DateTime end = start.AddMonths(1).AddHours(9);
            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(start, end);
            List<OperatorInputData> operatorInputDatas = await _operatorInputData.GetByDateTimeRangeAsync(start, end);
            var data = await CalDataAsync(3, start, end, sourceDatas, operatorInputDatas);
            if (data != null)
            {
                await _comToSiemens.AddAsync(data);
                await _reportUnitOfWork.SaveChangesAsync(); ;
            }
        }

        public async Task<ComToSiemens?> CalDataAsync(int saveType, DateTime startTime, DateTime endTime, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            ComToSiemens comToSiemens = new()
            {
                ReportedTime = DateTime.Now,
                Type = saveType,
                //PH = 0
            };
            //时间已经格式化过 不要再格式化
            //计算经过的天数
            int daysBetween = (endTime.Date - startTime.Date).Days;

            if (sourceDatas == null || sourceDatas.Count == 0)
                return null;
            if (operatorInputDatas == null || operatorInputDatas.Count == 0)
                return null;

            var WeekRangeData = DataToViewService.CalculateDailyProductionReportRange(startTime, endTime, sourceDatas, operatorInputDatas);
            List<decimal?> WeekRangeYield = WeekRangeData == null
                            ? new List<decimal?>() // 集合为null时返回空列表，避免空引用异常
                            : WeekRangeData.Select(report => report?.Cell2).ToList();
            var Materialcollection = new MaterialDataRangeCollection(startTime, daysBetween, WeekRangeYield, sourceDatas, operatorInputDatas);//获取7天的数据

            if (WeekRangeData == null) return comToSiemens;
            comToSiemens.Cell1 = (float?)MathTools.CalculateSum(WeekRangeData, x => x.Cell3);//产量累计
            comToSiemens.Cell2 = (float?)MathTools.CalculateSum(WeekRangeData, x => x.Cell2);//折百产量累计
            comToSiemens.Cell3 = (float)DataToViewService.CalculateRangeYield(Materialcollection);//时间范围内的收率
            comToSiemens.Cell4 = (float?)MathTools.CalculateFirstLastDifference(sourceDatas, x => x.Cell14);//气氨消耗
            comToSiemens.Cell5 = (float?)MathTools.CalculateFirstLastDifference(sourceDatas, x => x.Cell20);//羟基乙睛消耗
            comToSiemens.Cell6 = (float?)MathTools.CalculateFirstLastDifference(sourceDatas, x => x.Cell37);//稀硫酸消耗
            comToSiemens.Cell7 = (float?)MathTools.CalculateAverage(sourceDatas.Take(sourceDatas.Count - 1), x => x.Cell24);//摩尔比 做平均 (去掉最后一个值)

            comToSiemens.Cell8 = (float?)MathTools.CalculateAverage(operatorInputDatas.Take(operatorInputDatas.Count - 1), x => x.Cell21); //二乙睛含量平均(去掉最后一个值)
            comToSiemens.Cell9 = (float?)MathTools.CalculateAverage(operatorInputDatas.Take(operatorInputDatas.Count - 1), x => x.Cell23); //水分含量平均(去掉最后一个值)
            comToSiemens.Cell10 = (float?)MathTools.CalculateAverage(operatorInputDatas.Take(operatorInputDatas.Count - 1), x => x.Cell25); //未知物含量平均(去掉最后一个值)

            var usageWater = (float?)MathTools.CalculateFirstLastDifference(sourceDatas, x => x.Cell143);//水消耗 修改为sourcedata中143 读取 
            var usageElectric = (float?)MathTools.CalculateFirstLastDifference(operatorInputDatas, x => x.Cell73);//电消耗

            var lowPressData = (float?)MathTools.CalculateFirstLastDifference(operatorInputDatas, x => x.Cell71); //低压蒸汽消耗-手动录入
            var midellPressData = (float?)MathTools.CalculateFirstLastDifference(operatorInputDatas, x => x.Cell72); //中压蒸汽消耗-手动录入
            var usagesteam = lowPressData ?? 0 + midellPressData ?? 0;//蒸汽总消耗= 低压+中压//气消耗 修改为低压和中压相加

            comToSiemens.Cell11 = usageWater;
            comToSiemens.Cell12 = usageElectric;
            comToSiemens.Cell13 = usagesteam;
            //水电气后面的和前面重复
            comToSiemens.Cell48 = usageWater;
            comToSiemens.Cell49 = usageElectric;
            comToSiemens.Cell50 = usagesteam;
            float standUsageWater = (usageWater ?? 0f) * 0.1229f;
            float standUsageElectric = (usageElectric ?? 0f) * 128.6f;
            float standUsagesteam = usagesteam * 0.0857f;
            //单位产品耗能
            float standUsageTotal = standUsageWater + standUsageElectric + standUsagesteam;
            var allProduction = comToSiemens.Cell1;
            var standUsagePerProduct = allProduction != null && allProduction != 0 ? standUsageTotal / allProduction : 0f;
            comToSiemens.Cell44 = standUsagePerProduct;

            return comToSiemens;
        }
    }

}
