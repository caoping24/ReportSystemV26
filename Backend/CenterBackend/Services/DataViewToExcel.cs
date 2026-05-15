using CenterBackend.IServices;
using CenterBackend.Models;
using CenterBackend.Models.ExcelDataView;
using CenterReport.Repository.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Asn1.Ocsp;
using System.Globalization;

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
                else if (DataCollection.SheetType == SheetType.MonthReport && DataCollection is MonthWorkBook monthWorkBook)
                    writeResult = MonthWriteExcel(workbook, monthWorkBook);
                else if (DataCollection.SheetType == SheetType.YearReport && DataCollection is YearWorkBook yearWorkBook)
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
            ISheet srcSheet;
            //曲线表写入日期
            srcSheet = srcWorkbook.GetSheetAt(1);
            SetXlsxCellString(srcSheet,59,1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));
            srcSheet = srcWorkbook.GetSheetAt(3);
            SetXlsxCellString(srcSheet,59,1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));
            //白班
            var dataList = dayWorkBookData.DaySheet;
            if (dataList.Count == 0)
                return false;

            srcSheet = srcWorkbook.GetSheetAt(0);                                   //实际要写的表
            SetXlsxCellString(srcSheet, 51, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));    //记录日期
            //srcSheet.ForceFormulaRecalculation = false;                                 //关闭公式自动计算
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

            //夜班
            dataList = dayWorkBookData.NightSheet;
            if (dataList.Count == 0)
                return false;

            srcSheet = srcWorkbook.GetSheetAt(2);                                       //实际要写的表
            SetXlsxCellString(srcSheet, 51, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));    //记录日期
            //srcSheet.ForceFormulaRecalculation = false;                                 //关闭公式自动计算
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

            //考评表
            var shiftsList = dayWorkBookData.ShiftsAnalysis;
            if (shiftsList.Count == 0)
                return false;
            srcSheet = srcWorkbook.GetSheetAt(4);
            srcSheet.ForceFormulaRecalculation = true; 
            SetXlsxCellString(srcSheet,5, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));
            SetXlsxCellString(srcSheet,26, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));
            for (int i = 0; i < 2; i++)
            {
                if (i >= shiftsList.Count) break;
                var data = shiftsList[i];
                if (data == null) continue; // 如果 data 为空则跳过
                int targeRow = 21 * i;
                WriteDataRowsToExcel(data, srcSheet, 6 + targeRow, 2, 1, 33);
                WriteDataRowsToExcel(data, srcSheet, 11 + targeRow, 2, 34, 66);
                WriteDataRowsToExcel(data, srcSheet, 16 + targeRow, 2, 67, 99);
                WriteDataRowsToExcel(data, srcSheet, 21 + targeRow, 2, 100, 113);
            }
            srcSheet.GetRow(20).GetCell(2).SetCellValue("test");
            //日报表
            var dayList = dayWorkBookData.DayAnalysis;
            if (dayList == null)
                return false;
            srcSheet = srcWorkbook.GetSheetAt(5);                                       //实际要写的表
            SetXlsxCellString(srcSheet, 3, 3, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));  //记录日期
            SetXlsxCellValue(srcSheet, 5, 7, dayList.Cell1 ?? 0f);
            SetXlsxCellValue(srcSheet, 6, 7, dayList.Cell2 ?? 0f);
            SetXlsxCellValue(srcSheet, 7, 7, dayList.Cell3 ?? 0f);

            SetXlsxCellValue(srcSheet, 9, 7, dayList.Cell4 ?? 0f);
            SetXlsxCellValue(srcSheet, 10, 7, dayList.Cell5 ?? 0f);

            SetXlsxCellValue(srcSheet, 12, 7, dayList.Cell6 ?? 0f);

            SetXlsxCellValue(srcSheet, 21, 1, dayList.Cell7 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 3, dayList.Cell8 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 5, dayList.Cell9 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 7, dayList.Cell10 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 9, dayList.Cell11 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 11, dayList.Cell12 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 13, dayList.Cell13 ?? 0f);

            return true;
        }

        private static bool MonthWriteExcel(XSSFWorkbook srcWorkbook, MonthWorkBook monthWorkBookData)
        {
            ISheet srcSheet;
            var dayList = monthWorkBookData.MonthAnalysis;
            if (dayList == null)
                return false;
            srcSheet = srcWorkbook.GetSheetAt(0);                                       //实际要写的表
            SetXlsxCellString(srcSheet, 3, 3, monthWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));    //记录日期
            SetXlsxCellValue(srcSheet, 5, 7, dayList.Cell1 ?? 0f);
            SetXlsxCellValue(srcSheet, 6, 7, dayList.Cell2 ?? 0f);
            SetXlsxCellValue(srcSheet, 7, 7, dayList.Cell3 ?? 0f);

            SetXlsxCellValue(srcSheet, 9, 7, dayList.Cell4 ?? 0f);
            SetXlsxCellValue(srcSheet, 10, 7, dayList.Cell5 ?? 0f);

            SetXlsxCellValue(srcSheet, 12, 7, dayList.Cell6 ?? 0f);

            SetXlsxCellValue(srcSheet, 21, 1, dayList.Cell7 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 2, dayList.Cell8 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 3, dayList.Cell9 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 4, dayList.Cell10 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 5, dayList.Cell11 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 6, dayList.Cell12 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 7, dayList.Cell13 ?? 0f);
            return true;
        }
        private static bool YearWriteExcel(XSSFWorkbook srcWorkbook, YearWorkBook yearWorkBookData)
        {
            ISheet srcSheet;
            var dayList = yearWorkBookData.YearAnalysis;
            if (dayList == null)
                return false;
            srcSheet = srcWorkbook.GetSheetAt(0);                                       //实际要写的表
            SetXlsxCellString(srcSheet, 3, 3, yearWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));    //记录日期
            SetXlsxCellValue(srcSheet, 5, 7, dayList.Cell1 ?? 0f);
            SetXlsxCellValue(srcSheet, 6, 7, dayList.Cell2 ?? 0f);
            SetXlsxCellValue(srcSheet, 7, 7, dayList.Cell3 ?? 0f);

            SetXlsxCellValue(srcSheet, 9, 7, dayList.Cell4 ?? 0f);
            SetXlsxCellValue(srcSheet, 10, 7, dayList.Cell5 ?? 0f);

            SetXlsxCellValue(srcSheet, 12, 7, dayList.Cell6 ?? 0f);

            SetXlsxCellValue(srcSheet, 21, 1, dayList.Cell7 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 2, dayList.Cell8 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 3, dayList.Cell9 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 4, dayList.Cell10 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 5, dayList.Cell11 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 6, dayList.Cell12 ?? 0f);
            SetXlsxCellValue(srcSheet, 21, 7, dayList.Cell13 ?? 0f);
            return true;
        }
        private static bool WeekWriteExcel(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {


            WeekWriteExcelSheet1(srcWorkbook, weekWorkBookData);
            WeekWriteExcelSheet2(srcWorkbook, weekWorkBookData);
            WeekWriteExcelSheet3(srcWorkbook, weekWorkBookData);
            WeekWriteExcelSheet4(srcWorkbook, weekWorkBookData);
            WeekWriteExcelSheet5(srcWorkbook, weekWorkBookData);
            WeekWriteExcelSheet6(srcWorkbook, weekWorkBookData);
            WeekWriteExcelSheet7(srcWorkbook, weekWorkBookData);

            //WeekWriteExcelSheet13(srcWorkbook, weekWorkBookData);2026年5月14日 弃用
            return true;
        }

        private static void BatchWriteDataToExcel(SingleShift dataList, ISheet sheet, int rowIdx, int colIdx, int cellStart, int cellEnd)
        {
            if (dataList == null) return;
            int offset = 0;
            for (int i = cellStart; i <= cellEnd; i++)
            {
                offset++;
                var cellProperty = dataList.GetType().GetProperty($"Cell{i}");
                if (cellProperty == null) continue;

                var value = cellProperty.GetValue(dataList);
                if (value == null) continue;// 如果 data 为空则跳过
                if (value is float floatValue) // 检查 value 是否可以转换为 float
                {
                    SetXlsxCellValue(sheet, rowIdx, colIdx + offset, floatValue);
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

        private static void WriteDataRowsToExcel<T>(T dataItem, ISheet sheet, int rowIdx, int colIdx, int cellStart, int cellEnd)
        {
            if (dataItem == null) return;
            int offset = 0;
            for (int i = cellStart; i <= cellEnd; i++)
            {
                offset++;
                var cellProperty = dataItem.GetType().GetProperty($"Cell{i}");
                if (cellProperty == null) continue;

                var value = cellProperty.GetValue(dataItem);
                if (value == null) continue;// 如果 data 为空则跳过
                if (value is float floatValue) // 检查 value 是否可以转换为 float
                {
                    SetXlsxCellValue(sheet, rowIdx, colIdx + offset, floatValue);
                }
            }
        }
        private static void WriteDataColumnsToExcel<T>(T dataItem, ISheet sheet, int rowIdx, int colIdx, int cellStart, int cellEnd)
        {
            if (dataItem == null) return;
            int offset = 0;
            for (int i = cellStart; i <= cellEnd; i++)
            {
                offset++;
                var cellProperty = dataItem.GetType().GetProperty($"Cell{i}");
                if (cellProperty == null) continue;

                var value = cellProperty.GetValue(dataItem);
                if (value == null) continue;// 如果 data 为空则跳过
                if (value is float floatValue) // 检查 value 是否可以转换为 float
                {
                    SetXlsxCellValue(sheet, rowIdx + offset, colIdx, floatValue);
                }
            }
        }


        private static void WeekWriteExcelSheet1(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(0);
            srcSheet.ForceFormulaRecalculation = true;
            SetXlsxCellString(srcSheet, 2, 6, weekWorkBookData.ReportDate);
            SetXlsxCellString(srcSheet, 2, 7, weekWorkBookData.WeekNumberInYear);

            var dataList = weekWorkBookData.WorkSheet1;
            int i= 0;
            if (dataList != null && dataList.Count > i && dataList[i] != null)
            {
                var dataItem = dataList[i];
                WriteDataColumnsToExcel (dataItem, srcSheet, 3 + i, 6, 1, 28);
            }
        }

        private static void WeekWriteExcelSheet2(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(1);
            //srcSheet.ForceFormulaRecalculation = false;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet2;
            for (int i = 0; i < 7; i++)
            {
                if (dataList != null && dataList.Count > i && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");
                    SetXlsxCellString(srcSheet, 5 + i, 1, dateString);
                    WriteDataRowsToExcel(dataItem, srcSheet, 5 + i, 2, 1, 5);
                }
            }
        }
        private static void WeekWriteExcelSheet3(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(2);
            //srcSheet.ForceFormulaRecalculation = false;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet3;
            for (int i = 0; i < 7; i++)
            {
                if (dataList != null && dataList.Count > i && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");
                    SetXlsxCellString(srcSheet, 5 + i, 1, dateString);
                    WriteDataRowsToExcel(dataItem, srcSheet, 5 + i, 2, 1, 7);
                }
            }
        }
        private static void WeekWriteExcelSheet4(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(3);
            //srcSheet.ForceFormulaRecalculation = false;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet4;
            for (int i = 0; i < 7; i++)
            {
                if (dataList != null && dataList.Count > i && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");
                    SetXlsxCellString(srcSheet, 5 + i, 1, dateString);
                    WriteDataRowsToExcel(dataItem, srcSheet, 5 + i, 2, 1, 5);
                }
            }
        }
        private static void WeekWriteExcelSheet5(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(4);
            srcSheet.ForceFormulaRecalculation = true;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet5;
            for (int i = 0; i < 7; i++)
            {
                if (dataList != null && dataList.Count > i && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");
                    SetXlsxCellString(srcSheet, 5 + i, 1, dateString);
                    WriteDataRowsToExcel(dataItem, srcSheet, 5 + i, 2, 1, 5);
                }
            }
        }

        private static void WeekWriteExcelSheet6(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(5);
            //srcSheet.ForceFormulaRecalculation = false;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet6;
            for (int i = 0; i < 7; i++)
            {
                
                if (dataList != null && dataList.Count > i && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");
                    SetXlsxCellString(srcSheet, 6 + i, 1, dateString);
                    WriteDataRowsToExcel(dataItem, srcSheet, 6 + i, 2, 1, 14);
                }
            }
        }
        private static void WeekWriteExcelSheet7(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(6);
            //srcSheet.ForceFormulaRecalculation = false;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet7;
            var baseTime = weekWorkBookData.ReportedTime;
            for (int i = 0; i < 7; i++)
            {
                if (dataList != null && dataList.Count > i && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");
                    SetXlsxCellString(srcSheet, 5 + i, 1, dateString);
                    WriteDataRowsToExcel(dataItem, srcSheet, 5 + i, 2, 1, 3);
                }
            }
        }
       


        private static DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = (int)dt.DayOfWeek - (int)DayOfWeek.Monday;
            if (diff < 0) diff += 7;
            return dt.AddDays(-diff).Date;
        }
        private static float? CalculateAverage<T>(IEnumerable<T> data, Func<T, float?> selector)
        {
            if (data == null || !data.Any())// 空数据校验
                return null;
            var validValues = data.Select(selector).OfType<float>();
            float sum = 0f;
            int count = 0;
            foreach (var value in validValues)
            {
                sum += value;
                count++;
            }
            return count > 0 ? sum / count : (float?)null;
        }
    }
}
