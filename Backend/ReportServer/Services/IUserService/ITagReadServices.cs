using ReportServer.Models;

namespace ReportServer.Services.IUserService
{
    public interface ITagReadServices
    {
        Task<List<TagMap>> ReadAllTagsAsync();
        Task<TagMap?> ReadSingleTagsAsync(string tagName);
    }
}
