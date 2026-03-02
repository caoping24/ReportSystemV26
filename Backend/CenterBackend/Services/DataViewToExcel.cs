using CenterBackend.IServices;
using CenterBackend.Models;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace CenterBackend.Services
{
    public class DataViewToExcel : IDataViewToExcel
    {
        public async Task<bool> WriteXlsxAndSaveAsync<T>(T DataCollection) where T : BaseSheet
        {
            var modFilePath = DataCollection.ModFilePath;
            try
            {
                using var templateStream = new FileStream(modFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var workbook = new XSSFWorkbook(templateStream);
                bool writeResult = false;
                if (DataCollection.SheetType == SheetType.DayReport && DataCollection is DayWorkBook dayWorkBook)
                    writeResult = DayWriteExcel(workbook, dayWorkBook);
                else if (DataCollection.SheetType == SheetType.MonthReport && DataCollection is DayWorkBook monthWorkBook)
                    writeResult = MonthWriteExcel(workbook, monthWorkBook);
                else if (DataCollection.SheetType == SheetType.YearReport && DataCollection is DayWorkBook yearWorkBook)
                    writeResult = YearWriteExcel(workbook, yearWorkBook);
                else if (DataCollection.SheetType == SheetType.WeekReport && DataCollection is WeekWorkBook weekWorkBook)
                    writeResult = WeekWriteExcel(workbook, weekWorkBook);

                if (!writeResult)
                {
                    workbook.Close();
                    return false;
                }
                var fullPath = Path.Combine(DataCollection.Directory, DataCollection.FileName);
                using var outputStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write);// 保存文件到指定路径
                workbook.Write(outputStream);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // 写Xlsx数据
        private static bool DayWriteExcel(XSSFWorkbook srcWorkbook, DayWorkBook dayWorkBookData)
        {

            var dataList = dayWorkBookData.DaySheet;
            if (dataList.Count == 0)
                return false;
            ISheet srcSheet;

            //白班
            srcSheet = srcWorkbook.GetSheetAt(0);                                   //实际要写的表
            SetXlsxCellString(srcSheet, 51, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));    //记录日期
            srcSheet.ForceFormulaRecalculation = false;                                 //关闭公式自动计算
            for (int i = 0; i < 13; i++) 
            {
                if (i >= dataList.Count) break;
                var data = dataList[i];
                if (data == null) continue; // 如果 data 为空则跳过
                int colOffSet = 1;

                int targeRow = 5 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 1, 50 );
                targeRow = 21 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 51, 100);
                targeRow = 38 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 101, 150);
            }

            //夜班
            srcSheet = srcWorkbook.GetSheetAt(2);                                       //实际要写的表
            SetXlsxCellString(srcSheet, 51, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));    //记录日期
            srcSheet.ForceFormulaRecalculation = false;                                 //关闭公式自动计算
            for (int i = 0; i < 13; i++)
            {
                if (i >= dataList.Count) break;
                var data = dataList[i];
                if (data == null) continue; // 如果 data 为空则跳过
                int colOffSet = 1;

                int targeRow = 5 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 1, 50);
                targeRow = 21 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 51, 100);
                targeRow = 38 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 101, 150);
            }
            return true;
        }

        private static bool MonthWriteExcel(XSSFWorkbook srcWorkbook, DayWorkBook monthWorkBookData)
        {
            return true;
        }

        private static bool YearWriteExcel(XSSFWorkbook srcWorkbook, DayWorkBook yearWorkBookData)
        {
            return true;
        }
        private static bool WeekWriteExcel(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            return true;
        }


        private static void BatchWriteDataToExcel(DayWorkSheet dataList, ISheet sheet, int rowIdx, int colIdx, int cellStart, int cellEnd)
        {
            if (dataList == null) return;
            for (int i = cellStart; i <= cellEnd; i++)
            {
                var cellProperty = dataList.GetType().GetProperty($"Cell{i}");
                if (cellProperty == null) continue;

                var value = cellProperty.GetValue(dataList);
                if (value == null) continue;// 如果 data 为空则跳过
                if (value is float floatValue) // 检查 value 是否可以转换为 float
                {
                    SetXlsxCellValue(sheet, rowIdx, colIdx + i, floatValue);
                }
            }
        }

        private static void SetXlsxCellValue(ISheet sheet, int rowIdx, int colIdx, float value)
        {
            IRow row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);// 获取或创建行
            ICell Cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);// 获取或创建单元格
            Cell.SetCellValue(value);// 赋值

        }
        private static void SetXlsxCellString(ISheet sheet, int rowIdx, int colIdx, string? value)
        {
            if (string.IsNullOrEmpty(value))// 如果值为空或null，则不写入单元格
                return;
            IRow row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);// 获取或创建行
            ICell Cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);// 获取或创建单元格
            Cell.SetCellValue(value);// 赋值
        }

    }
}
