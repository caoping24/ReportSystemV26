using CenterBackend.IServices;
using CenterBackend.Models.CalculateData;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using CenterReport.Repository.Services;
using CenterUser.Repository;
using Hangfire.Common;
using MathNet.Numerics.LinearAlgebra.Factorization;
using Microsoft.EntityFrameworkCore.Diagnostics;
using NPOI.SS.Formula.Functions;
using Org.BouncyCastle.Asn1.X509;


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

            DateTime end = DateTime.Now.Date.AddHours(8);
            DateTime start = end.AddDays(-1);
    
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

            DateTime end = DateTime.Now.Date.AddHours(8);
            DateTime start = end.AddDays(-7);

            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(start, end);
            List<OperatorInputData> operatorInputDatas = await _operatorInputData.GetByDateTimeRangeAsync(start, end);

            var data = await  CalDataAsync(2, start, end, sourceDatas, operatorInputDatas);
            if (data != null)
            {
                await _comToSiemens.AddAsync(data);
                await _reportUnitOfWork.SaveChangesAsync(); ;
            }
        }
        public async Task MonthlyDay1_0830()
        {

            DateTime end = DateTime.Now.Date.AddHours(8);
            DateTime start = new(end.Year, end.Month -1 , 1, 8, 0, 0);//上个月1号8点

            List<SourceData> sourceDatas = await _sourceData.GetByDateTimeRangeAsync(start, end);
            List<OperatorInputData> operatorInputDatas = await _operatorInputData.GetByDateTimeRangeAsync(start, end);

            var  data = await CalDataAsync(3, start, end, sourceDatas, operatorInputDatas);
            if (data != null)
            {
                await _comToSiemens.AddAsync(data);
                await _reportUnitOfWork.SaveChangesAsync(); ;
            }
        }

        public async Task<ComToSiemens?> CalDataAsync(int saveType, DateTime start, DateTime end, List<SourceData> sourceDatas, List<OperatorInputData> operatorInputDatas)
        {
            ComToSiemens comToSiemens = new()
            { 
                    ReportedTime = DateTime.Now,
                    Type = saveType,
                    //PH = 0
            };

            DateTime startTime = start.Date.AddHours(8);
            DateTime endTime = end.Date.AddHours(8);

            if (sourceDatas == null || sourceDatas.Count == 0)
                return null;
            if (operatorInputDatas == null || operatorInputDatas.Count == 0)
                return null;
            //需要计算的数据
            //List<ProductionDataCollection> productionDataCollections = DataToViewService.CalculateForSheet3TimeRange(startTime, endTime, operatorInputDatas);//计算产品数据
            
            //var allProduction  = DataToViewService.CalculateSum(productionDataCollections, x => x.TotalResult.AllProduction); //产量累计
            //comToSiemens.Cell1 = allProduction;
            //comToSiemens.Cell2 = DataToViewService.CalculateAverage(productionDataCollections, x => x.TotalResult.AllYield);//折百平均
            //comToSiemens.Cell3 = DataToViewService.CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_1);//收率平均

            //comToSiemens.Cell4 = DataToViewService.CalculateFirstLastDifference(sourceDatas, x => x.Cell14);//气氨消耗
            //comToSiemens.Cell5 = DataToViewService.CalculateFirstLastDifference(sourceDatas, x => x.Cell20);//羟基乙睛消耗
            //comToSiemens.Cell6 = DataToViewService.CalculateFirstLastDifference(sourceDatas, x => x.Cell37);//稀硫酸消耗
            //comToSiemens.Cell7 = DataToViewService.CalculateAverage(sourceDatas, x => x.Cell24);//摩尔比 做平均

            //comToSiemens.Cell8   = DataToViewService.CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_1); //二乙睛含量平均
            //comToSiemens.Cell9 = DataToViewService.CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_3); //水分含量平均
            //comToSiemens.Cell10 = DataToViewService.CalculateAverage(productionDataCollections, x => x.TotalResult.AllAverage_5); //未知物含量平均

            var usageWater = DataToViewService.CalculateFirstLastDifference(sourceDatas, x => x.Cell143);//水消耗 修改为sourcedata中143 读取 
            var usageElectric = DataToViewService.CalculateFirstLastDifference(operatorInputDatas, x => x.Cell73);//电消耗


            var lowPressData = DataToViewService.CalculateFirstLastDifference(operatorInputDatas, x => x.Cell71); //低压蒸汽消耗-手动录入
            var midellPressData = DataToViewService.CalculateFirstLastDifference(operatorInputDatas, x => x.Cell72); //中压蒸汽消耗-手动录入
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

            float standUsageTotal = standUsageWater + standUsageElectric + standUsagesteam;
            //var standUsagePerProduct = allProduction != null && allProduction != 0 ? standUsageTotal / allProduction : 0f;

            //单位产品耗能
            //comToSiemens.Cell44 = standUsagePerProduct;

            return comToSiemens;
        }
    }

}
