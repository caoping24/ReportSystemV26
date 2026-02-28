using CenterBackend.IServices;
using CenterBackend.Models;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace CenterBackend.Services
{
    public class DataViewToExcel : IDataViewToExcel
    {

        public async Task<bool> DayWriteAndSaveAsync(PathAndName fileInfo, DayReportCollection data)
        {
            if (fileInfo.Type == -1)//路径为空
                return false;
            if (data == null || data.ReportList == null || data.ReportList.Count == 0)
                return false;
            try
            {
                using var templateStream = new FileStream(fileInfo.FullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var workbook = new XSSFWorkbook(templateStream);

                switch (fileInfo.Type)
                {
                    case 1 :
                        DayWriteExcel( workbook,  data);
                        break;
                    case 2:
                        MonthWriteExcel(workbook, data);
                        break;
                    case 3:
                        DayWriteExcel(workbook, data);
                        break;
                    case 4:
                        DayWriteExcel(workbook, data);
                        break;
                    default:
                        return false;
                }
                using var outputStream = new FileStream(fileInfo.FullPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await Task.Run(() => workbook.Write(outputStream)); // 异步写入，符合async规范
                await outputStream.FlushAsync(); // 强制刷新缓冲区，确保数据落盘
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void DayWriteExcel(XSSFWorkbook workbook, DayReportCollection data)
        {
            var sheet = workbook.GetSheetAt(0);
            int startRow = 2; // 假设模板表头在第1行，数据从第2行开始
            foreach (var report in data.ReportList)
            {
                var row = sheet.GetRow(startRow) ?? sheet.CreateRow(startRow);
                startRow++;

                row.CreateCell(0).SetCellValue(report.TimePoint);
            }
        }
        private static void MonthWriteExcel(XSSFWorkbook workbook, DayReportCollection data)
        {
            var sheet = workbook.GetSheetAt(0);
            int startRow = 2; // 假设模板表头在第1行，数据从第2行开始
            foreach (var report in data.ReportList)
            {
                var row = sheet.GetRow(startRow) ?? sheet.CreateRow(startRow);
                startRow++;

                row.CreateCell(0).SetCellValue(report.TimePoint);
            }
        }


    }
}
