namespace ReportServer.Services
{
    namespace IUserService
    {
        public interface ICollectWinccDatas
        {
            Task<bool> ReadAndSaveDataAsync();
        }
    }
}
