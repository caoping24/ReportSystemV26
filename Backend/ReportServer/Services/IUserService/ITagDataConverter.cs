using CenterReport.Repository.Models;
using ReportServer.Models;

namespace ReportServer.Services.IUserService
{
    public interface ITagDataConverter
    {
        SourceData? ConvertTagsToSourceData(List<TagMap>? tags);
    }
}
