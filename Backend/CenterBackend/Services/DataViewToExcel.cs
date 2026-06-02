using CenterBackend.IServices;
using CenterBackend.Models.ExcelDataView;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Reflection;

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
        // 写报表数据
        private static bool DayWriteExcel(XSSFWorkbook srcWorkbook, DayWorkBook dayWorkBookData)
        {
            ISheet srcSheet;
            //曲线表写入日期
            srcSheet = srcWorkbook.GetSheetAt(1);
            SetXlsxCellString(srcSheet, 59, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));
            srcSheet = srcWorkbook.GetSheetAt(3);
            SetXlsxCellString(srcSheet, 59, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));
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
                int colOffSet = 2;

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
                int colOffSet = 2;

                int targeRow = 5 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 1, 50);
                targeRow = 21 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 51, 100);
                targeRow = 38 + i;
                BatchWriteDataToExcel(data, srcSheet, targeRow, colOffSet, 101, 150);
            }

            srcSheet = srcWorkbook.GetSheetAt(4);
            //srcSheet.ForceFormulaRecalculation = true;
            var target = dayWorkBookData.ShiftsAnalysis; if (target == null) return false;
            SetXlsxCellString(srcSheet, 2, 1, dayWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));
            SetXlsxCellString(srcSheet, 5, 1, target.TimePoint.ToString("yyyy-MM-dd"));//写日期
            if (target.Data != null && target.Data.ShiftBatches.Count > 0)
            {
                WriteDataRowsToExcel(target.Data, srcSheet, 5, 16, 2, 3);
                for (var j = 0; j < target.Data.MaxBatches; j++)
                {
                    var item = target.Data.ShiftBatches[j];
                    WriteDataRowsToExcel(item, srcSheet, j * 3 + 5, 9, 4, 5);
                    for (var i = 0; i < item.MaxBatches; i++)
                    {
                        if (i < item.Batches.Count)
                        {
                            if (item.Batches[i].IsEmpty == false)
                            {
                                WriteDataRowsToExcel(item.Batches[i], srcSheet, j * 3 + i + 5, 2, 1, 3);
                            }
                        }
                    }
                }
            }
            //日报表
            srcSheet = srcWorkbook.GetSheetAt(5); //实际要写的表
            srcSheet.ForceFormulaRecalculation = true;
            //SetXlsxCellString(srcSheet, 2, 6, dayList.ReportDate);
            //SetXlsxCellString(srcSheet, 2, 7, dayList.WeekNumberInYear);
            var dayResult = dayWorkBookData.DayAnalysis;
            int k = 0;
            if (dayResult != null)
            {
                WriteDataColumnsToExcel(dayResult, srcSheet, 3 + k, 6, 1, 28);
            }
            return true;
        }
        private static bool MonthWriteExcel(XSSFWorkbook srcWorkbook, MonthWorkBook monthWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(0);
            srcSheet.ForceFormulaRecalculation = true;
            SetXlsxCellString(srcSheet, 2, 6, monthWorkBookData.ReportDate);
            SetXlsxCellString(srcSheet, 2, 7, monthWorkBookData.WeekNumberInYear);
            var dataList = monthWorkBookData.MonthAnalysis;
            int i = 0;
            if (dataList != null && dataList.Count > i && dataList[i] != null)
            {
                var dataItem = dataList[i];
                WriteDataColumnsToExcel(dataItem, srcSheet, 3 + i, 6, 1, 28);
            }
            return true;
        }
        private static bool YearWriteExcel(XSSFWorkbook srcWorkbook, YearWorkBook yearWorkBookData)
        {
            //ISheet srcSheet;
            //var dayList = yearWorkBookData.YearAnalysis;
            //if (dayList == null)
            //    return false;
            //srcSheet = srcWorkbook.GetSheetAt(0);                                       //实际要写的表
            //SetXlsxCellString(srcSheet, 3, 3, yearWorkBookData.ReportedTime.ToString("yyyy-MM-dd"));    //记录日期
            //SetXlsxCellValue(srcSheet, 5, 7, dayList.Cell1 ?? 0f);
            //SetXlsxCellValue(srcSheet, 6, 7, dayList.Cell2 ?? 0f);
            //SetXlsxCellValue(srcSheet, 7, 7, dayList.Cell3 ?? 0f);

            //SetXlsxCellValue(srcSheet, 9, 7, dayList.Cell4 ?? 0f);
            //SetXlsxCellValue(srcSheet, 10, 7, dayList.Cell5 ?? 0f);

            //SetXlsxCellValue(srcSheet, 12, 7, dayList.Cell6 ?? 0f);

            //SetXlsxCellValue(srcSheet, 21, 1, dayList.Cell7 ?? 0f);
            //SetXlsxCellValue(srcSheet, 21, 2, dayList.Cell8 ?? 0f);
            //SetXlsxCellValue(srcSheet, 21, 3, dayList.Cell9 ?? 0f);
            //SetXlsxCellValue(srcSheet, 21, 4, dayList.Cell10 ?? 0f);
            //SetXlsxCellValue(srcSheet, 21, 5, dayList.Cell11 ?? 0f);
            //SetXlsxCellValue(srcSheet, 21, 6, dayList.Cell12 ?? 0f);
            //SetXlsxCellValue(srcSheet, 21, 7, dayList.Cell13 ?? 0f);
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
            WeekWriteExcelSheet8(srcWorkbook, weekWorkBookData);

            WeekWriteExcelSheet9(srcWorkbook, weekWorkBookData);//图表数据

            return true;
        }
        //写单个sheet
        private static void WeekWriteExcelSheet1(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(0);
            srcSheet.ForceFormulaRecalculation = true;
            SetXlsxCellString(srcSheet, 2, 6, weekWorkBookData.ReportDate);
            SetXlsxCellString(srcSheet, 2, 7, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet1;
            int i = 0;
            if (dataList != null && dataList.Count > i && dataList[i] != null)
            {
                var dataItem = dataList[i];
                WriteDataColumnsToExcel(dataItem, srcSheet, 3 + i, 6, 1, 28);
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
            srcSheet.ForceFormulaRecalculation = true;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet4;
            for (int i = 0; i < 7; i++)
            {
                if (dataList != null && i < dataList.Count && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");//写日期
                    SetXlsxCellString(srcSheet, 5 + i * 6, 1, dateString);
                    WriteDataRowsToExcel(dataItem.Data, srcSheet, i * 6 + 5, 16, 2, 3);
                    if (dataItem.Data != null && dataItem.Data.ShiftBatches.Count > 0)
                    {

                        for (var j = 0; j < dataItem.Data.MaxBatches; j++)
                        {
                            var item = dataItem.Data.ShiftBatches[j];
                            for (var k = 0; k < item.MaxBatches; k++)
                            {
                                if (k < item.Batches.Count)
                                {
                                    if (item.Batches[k].IsEmpty == false)
                                    {
                                        WriteDataRowsToExcel(item.Batches[k], srcSheet, i * 6 + j * 3 + k + 5, 2, 1, 3);
                                    }
                                }
                                if (k == 0)
                                    WriteDataRowsToExcel(item, srcSheet, i * 6 + j * 3 + k + 5, 9, 4, 5);
                            }
                        }
                    }
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
                    WriteDataRowsToExcel(dataItem, srcSheet, 6 + i, 2, 1, 13);
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
        private static void WeekWriteExcelSheet8(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(7);
            //srcSheet.ForceFormulaRecalculation = false;
            SetXlsxCellString(srcSheet, 2, 1, weekWorkBookData.WeekNumberInYear);
            var dataList = weekWorkBookData.WorkSheet8;
            for (int i = 0; i < 7; i++)
            {

                if (dataList != null && dataList.Count > i && dataList[i] != null)
                {
                    var dataItem = dataList[i];
                    var dateString = dataList[i].TimePoint.ToString("yyyy-MM-dd");
                    SetXlsxCellString(srcSheet, 6 + i, 1, dateString);
                    WriteDataRowsToExcel(dataItem, srcSheet, 6 + i, 2, 1, 7);
                }
            }
        }
        private static void WeekWriteExcelSheet9(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData)
        {
            ISheet srcSheet;
            srcSheet = srcWorkbook.GetSheetAt(8);

            int[] indexInMaterialDatas = new[] { 3, 4, 6, 7, 8, 9, 10, 11, 12, 13, 15 };
            string[] chartNames = [
                                    "羟基浓度(g/l)",
                                    "氨腈摩尔比",
                                    "反应压力(MPa)",
                                    "羟基加热温度(°C)",
                                    "氨汽混合温度(°C)",
                                    "管反热点温度(°C)",
                                    "预冷器结晶温度(°C)",
                                    "一次结晶温度(°C)",
                                    "降膜蒸发温度(°C)",
                                    "二次结晶温度(°C)" ,
                                    "废液排放(m³/t)"];
            int[,,] address = new int[,,]
                                    {
                                        { {3}  , {3} ,},
                                        { {53} , {3} ,},
                                        { {103}, {3} ,},
                                        { {153}, {3} ,},
                                        { {203}, {3} ,},
                                        { {253}, {3} ,},
                                        { {303}, {3} ,},
                                        { {353}, {3} ,},
                                        { {403}, {3} ,},
                                        { {453}, {3} ,},
                                        { {503}, {3} ,},
                                    };
            for (int i = 0; i < 11; i++)
            {
                int index = indexInMaterialDatas[i];
                int row = address[i, 0, 0];
                int col = address[i, 1, 0];
                string chartName = chartNames[i];
                WeekWriteExcelChart(srcSheet, weekWorkBookData, index, row, col, chartName);
            }

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="srcSheet">表引用</param>
        /// <param name="weekWorkBookData">数据引用</param>
        /// <param name="dataIndex">数据在MaterialDatas中的index</param>
        /// <param name="rowStart"></param>
        /// <param name="cloumnStart"></param>
        private static void WeekWriteExcelChart(ISheet srcSheet,
                                                WeekWorkBook weekWorkBookData,
                                                int dataIndex,
                                                int rowStart,
                                                int cloumnStart,
                                                string chartName)
        {
            var dataList = weekWorkBookData.WorkSheet9;
            List<decimal> dataItem;
            if (dataList != null)
            {
                dataItem = dataList.DailyCollections
                                .Select(daily => daily.MaterialDatas.FirstOrDefault(m => m.Index == dataIndex)? //第一个
                                .Specific ?? 0)
                                .ToList();

                SetXlsxCellString(srcSheet, rowStart - 2, cloumnStart - 1, chartName ?? "曲线");
                for (int i = 0; i < 7; i++)
                {
                    var temp = dataItem[i];
                    SetXlsxCellValue(srcSheet, i + rowStart, cloumnStart, temp);
                }

            }
        }

        //写excel方法
        private static void BatchWriteDataToExcel(SingleShift dataList, ISheet sheet, int rowIdx, int colIdx, int cellStart, int cellEnd)
        {
            if (dataList == null) return;
            var properties = new PropertyInfo[cellEnd - cellStart + 1];
            var type = typeof(SingleShift);
            for (int i = cellStart; i <= cellEnd; i++)
            {
                var prop = type.GetProperty($"Cell{i}");
                if (prop == null) continue;
                properties[i - cellStart] = prop;
            }

            for (int idx = 0; idx < properties.Length; idx++)
            {
                var prop = properties[idx];
                if (prop == null) continue;
                var rawValue = prop.GetValue(dataList);
                if (rawValue == null) continue;
                decimal value = Convert.ToDecimal(rawValue);
                SetXlsxCellValue(sheet, rowIdx, colIdx + idx, value);
            }

            //if (dataList == null) return;
            //int offset = 0;
            //for (int i = cellStart; i <= cellEnd; i++)
            //{
            //    offset++;
            //    var cellProperty = dataList.GetType().GetProperty($"Cell{i}");
            //    if (cellProperty == null) continue;

            //    decimal? value = (decimal?) cellProperty.GetValue(dataList);
            //    if (value == null) continue;// 如果 data 为空则跳过
            //    SetXlsxCellValue(sheet, rowIdx, colIdx + offset, value.Value);
            //}
        }
        private static void SetXlsxCellValue(ISheet sheet, int rowIdx, int colIdx, decimal value)
        {

            IRow row = sheet.GetRow(rowIdx) ?? sheet.CreateRow(rowIdx);// 获取或创建行
            ICell Cell = row.GetCell(colIdx) ?? row.CreateCell(colIdx);// 获取或创建单元格
            Cell.SetCellValue((double)value);// 赋值

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

                decimal? value = (decimal?)cellProperty.GetValue(dataItem);
                if (value == null) continue;// 如果 data 为空则跳过
                SetXlsxCellValue(sheet, rowIdx, colIdx + offset, value.Value);
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

                decimal? value = (decimal?)cellProperty.GetValue(dataItem);
                if (value == null) continue;// 如果 data 为空则跳过
                SetXlsxCellValue(sheet, rowIdx + offset, colIdx, value.Value);
            }
        }
        //图表的操作
        private static XSSFChart? GetChartByObjectName(ISheet sheet, string chartObjectName)//根据对象名称查找指定图表
        {
            if (string.IsNullOrWhiteSpace(chartObjectName))
                return null;
            if (sheet is not XSSFSheet xssfSheet)
                throw new NotSupportedException("仅支持 .xlsx 格式文件");

            XSSFDrawing? drawing = xssfSheet.GetDrawingPatriarch() as XSSFDrawing;
            if (drawing == null || drawing.Count() == 0)
                return null;

            foreach (XSSFShape shape in drawing.Cast<XSSFShape>())
            {
                if (shape is XSSFGraphicFrame graphicFrame &&
                    string.Equals(graphicFrame.Name, chartObjectName, StringComparison.OrdinalIgnoreCase))
                {
                    XSSFChart? chart = GetChartFromGraphicFrame(drawing, graphicFrame);
                    if (chart != null)
                        return chart;
                }
            }
            return null;
        }
        private static XSSFChart? GetChartFromGraphicFrame(XSSFDrawing drawing, XSSFGraphicFrame targetFrame)
        {
            foreach (var relationPart in drawing.RelationParts)
            {
                if (relationPart.DocumentPart is XSSFChart chart)
                {
                    try
                    {
                        XSSFGraphicFrame? graphicFrame = chart.GetGraphicFrame();
                        if (graphicFrame != null && graphicFrame == targetFrame)
                        {
                            return chart;
                        }
                    }
                    catch (NotImplementedException)
                    {
                        return null;
                    }
                }
            }
            return null;
        }
    }
}
