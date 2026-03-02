using CenterBackend.Models.ExcelDataView;

namespace CenterBackend.IServices
{
    public interface IDataViewToExcel
    {
        Task<bool> WriteXlsxAndSaveAsync<T>(T DataCollection) where T : BaseSheet;
    }
}
