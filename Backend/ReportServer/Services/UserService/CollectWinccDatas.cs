using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using ReportServer.Models;
using ReportServer.Services.IUserService;
using static FastExpressionCompiler.ExpressionCompiler;
using static ReportServer.Services.UserService.LogServices;

namespace ReportServer.Services.UserService
{
    public class CollectWinccDatas : ICollectWinccDatas
    {
        private readonly ITagReadServices _tagReadServices;
        private readonly ITagDataConverter _tagDataConverter;
        private readonly IReportRepository<SourceData> _sourceData;
        private readonly IReportRepository<ReportRecord> _reportRecord;
        private readonly IReportUnitOfWork _reportUnitOfWork;
        public CollectWinccDatas(ITagReadServices tagReadServices,
                                ITagDataConverter tagDataConverter, 
                                IReportRepository<SourceData> sourceData,
                                IReportUnitOfWork reportUnitOfWork, 
                                IReportRepository<ReportRecord> reportRecord)
        {
            _tagReadServices = tagReadServices;
            _tagDataConverter = tagDataConverter;
            _sourceData = sourceData;
            _reportUnitOfWork = reportUnitOfWork;
            _reportRecord = reportRecord;
        }
        public async Task<bool> ReadAndSaveDataAsync()
        {
            await AutoAddRecords();//固定时间向AutoAddRecords添加记录
            try
            {
                List<TagMap>? tags = await _tagReadServices.ReadAllTagsAsync();

                SourceData? result = _tagDataConverter.ConvertTagsToSourceData(tags);
                if (result == null)
                {
                    await AsyncLogHelper.LogErrorAsync("标签列表为空,数据收集失败.");
                    return false;
                }
                await _sourceData.AddAsync(result);
                await _reportUnitOfWork.SaveChangesAsync();
                await AsyncLogHelper.LogInfoAsync("数据收集成功" + DateTime.Now.ToString());
                return true;
            }
            catch (Exception ex)
            {
                await AsyncLogHelper.LogErrorAsync($"数据收集失败.:{ ex}");
                return false;
            }
        }
        private async Task AutoAddRecords()
        {
            //DateTime now = DateTime.Now;
            DateTime now = new DateTime(2025, 12,29,9,0,1);
            List<ReportRecord> reportRecordList = [];

            if (IsDaily(now))
            {
                reportRecordList.Add(new ReportRecord
                {
                    ReportedTime = now,
                    Type = 1,
                    //Description = ""
                });
            }
            if (IsMonthly(now))
            {
                reportRecordList.Add(new ReportRecord
                {
                    ReportedTime = now,
                    Type = 2,
                    //Description = ""
                });
            }
            if (IsYearly(now))
            {
                reportRecordList.Add(new ReportRecord
                {
                    ReportedTime = now,
                    Type = 3,
                    //Description = ""
                });
            }
            if (IsWeekly(now))
            {
                reportRecordList.Add(new ReportRecord
                {
                    ReportedTime = now,
                    Type = 4,
                    Description = GetYearWeekString(now)
                });
            }
            try
            {
                if (reportRecordList.Count > 0)
                {
                    foreach (var item in reportRecordList)
                    {
                        await _reportRecord.AddAsync(item);
                    }
                    await _reportUnitOfWork.SaveChangesAsync();
                }
            }
            catch (Exception ex)
            {
                await AsyncLogHelper.LogErrorAsync($"ReportRecord添加失败.ex:{ex}");
            }
        }

        // 判断时间=9点
        private static bool IsDaily(DateTime time) => time.Hour == 9;
        //判断时间=1号9点
        private static bool IsMonthly(DateTime time) => time.Day == 1 && time.Hour == 9;
        // 判断时间=1月1号9点
        private static bool IsYearly(DateTime time) => time.Month == 1 && time.Day == 1 && time.Hour == 9;
        // 判断时间=周一9点
        private static bool IsWeekly(DateTime time) => time.DayOfWeek == DayOfWeek.Monday && time.Hour == 9;
        // 获取日期所在周数
        private static string GetYearWeekString(DateTime now)
        {
            DayOfWeek day = now.DayOfWeek;
            int days = day - DayOfWeek.Monday;
            if (days < 0) days += 7;

            DateTime monday = now.AddDays(-days);// 取当前周一
            var culture = System.Globalization.CultureInfo.InvariantCulture;// 计算 ISO 标准周数
            var calendar = culture.Calendar;
            int weekNum = calendar.GetWeekOfYear(
                monday,
                System.Globalization.CalendarWeekRule.FirstFourDayWeek,
                DayOfWeek.Monday
            );
            int year = monday.Year;
            return $"{year}年第{weekNum}周";
        }
    }

}
