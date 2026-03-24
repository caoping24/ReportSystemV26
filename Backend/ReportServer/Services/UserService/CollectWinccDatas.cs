using CenterReport.Repository.IServices;
using CenterReport.Repository.Models;
using ReportServer.Models;
using ReportServer.Services.IUserService;
using static ReportServer.Services.UserService.LogServices;

namespace ReportServer.Services.UserService
{
    public class CollectWinccDatas : ICollectWinccDatas
    {
        private readonly ITagReadServices _tagReadServices;
        private readonly ITagDataConverter _tagDataConverter;
        private readonly IReportRepository<SourceData> _sourceData;
        private readonly IReportUnitOfWork _reportUnitOfWork;
        public CollectWinccDatas(ITagReadServices tagReadServices, ITagDataConverter tagDataConverter, IReportRepository<SourceData> sourceData, IReportUnitOfWork reportUnitOfWork)
        {
            _tagReadServices = tagReadServices;
            _tagDataConverter = tagDataConverter;
            _sourceData = sourceData;
            _reportUnitOfWork = reportUnitOfWork;
        }
        public async Task<bool> ReadAndSaveDataAsync()
        {
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

    }

}
