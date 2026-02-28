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
            string fullPath = Path.Combine(DataCollection.Directory, DataCollection.FileName);
            try
            {
                using var templateStream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var workbook = new XSSFWorkbook(templateStream);
                var reportedTime = DataCollection.ReportedTime;
                switch (DataCollection.SheetType)
                {
                    case SheetType.DayReport:
                        if (DataCollection is DayWorkBook dayWorkBook)
                            return DayWriteExcel(workbook, dayWorkBook, reportedTime);
                        break;
                    case SheetType.MonthReport:
                        if (DataCollection is DayWorkBook monthWorkBook)
                            return MonthWriteExcel(workbook, monthWorkBook, reportedTime);
                        break;
                    case SheetType.YearReport:
                        if (DataCollection is DayWorkBook yearWorkBook)
                            return YearWriteExcel(workbook, yearWorkBook, reportedTime);
                        break;
                    case SheetType.WeekReport:
                        if (DataCollection is WeekWorkBook weekWorkBook)
                            return WeekWriteExcel(workbook, weekWorkBook, reportedTime);
                        break;
                    case SheetType.OtherReport:
                        return false;
                    default:
                        return false;
                }
                using var outputStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await Task.Run(() => workbook.Write(outputStream)); // 异步写入，符合async规范
                await outputStream.FlushAsync(); // 强制刷新缓冲区，确保数据落盘
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        /// <summary>
        /// 写Xlsx数据  白班
        /// </summary>
        private static bool DayWriteExcel(XSSFWorkbook srcWorkbook, DayWorkBook dayWorkBookData, string reportedTime)
        {
            ISheet srcSheet = srcWorkbook.GetSheetAt(0);        //实际要写的表
            SetXlsxCellString(srcSheet, 51, 1, reportedTime);   //记录日期
            srcSheet.ForceFormulaRecalculation = false;         //关闭公式自动计算
            var dataList = dayWorkBookData.DaySheet;
            for (int i = 0; i < 13; i++)
            {
                var data = dataList.ElementAt(i);
                if (data == null) continue; // 如果 data 为空则跳过

                int Range1 = 5 + i;
                int Range2 = 21 + i;
                int Range3 = 38 + i;

                // 从Excel第2列开始写入
                //Rang1 
                if (data.Cell1 != null) { SetXlsxCellValue(srcSheet, Range1, 2, data.Cell1.Value); }
                if (data.Cell2 != null) { SetXlsxCellValue(srcSheet, Range1, 3, data.Cell2.Value); }
                if (data.Cell3 != null) { SetXlsxCellValue(srcSheet, Range1, 4, data.Cell3.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell4 != null && prevData != null && prevData != null && prevData.Cell4 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell4);
                        float prevVal = Convert.ToSingle(prevData.Cell4);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 5, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 5, 0); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell5 != null && prevData != null && prevData.Cell5 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell5);
                        float prevVal = Convert.ToSingle(prevData.Cell5);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 6, 0); }
                if (data.Cell6 != null) { SetXlsxCellValue(srcSheet, Range1, 7, data.Cell6.Value); }
                if (data.Cell7 != null) { SetXlsxCellValue(srcSheet, Range1, 8, data.Cell7.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell8 != null && prevData != null && prevData.Cell8 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell8);
                        float prevVal = Convert.ToSingle(prevData.Cell8);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 9, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 9, 0); }
                if (data.Cell9 != null) { SetXlsxCellValue(srcSheet, Range1, 10, data.Cell9.Value); }
                if (data.Cell10 != null) { SetXlsxCellValue(srcSheet, Range1, 11, data.Cell10.Value); }
                if (data.Cell11 != null) { SetXlsxCellValue(srcSheet, Range1, 12, data.Cell11.Value); }
                if (data.Cell12 != null) { SetXlsxCellValue(srcSheet, Range1, 13, data.Cell12.Value); }
                if (data.Cell13 != null) { SetXlsxCellValue(srcSheet, Range1, 14, data.Cell13.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell14 != null && prevData != null && prevData.Cell14 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell14);
                        float prevVal = Convert.ToSingle(prevData.Cell14);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 15, 0); }
                if (data.Cell15 != null) { SetXlsxCellValue(srcSheet, Range1, 16, data.Cell15.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell16 != null && prevData != null && prevData.Cell16 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell16);
                        float prevVal = Convert.ToSingle(prevData.Cell16);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 17, 0); }
                if (data.Cell17 != null) { SetXlsxCellValue(srcSheet, Range1, 18, data.Cell17.Value); }
                if (data.Cell18 != null) { SetXlsxCellValue(srcSheet, Range1, 19, data.Cell18.Value); }
                if (data.Cell19 != null) { SetXlsxCellValue(srcSheet, Range1, 20, data.Cell19.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell20 != null && prevData != null && prevData.Cell20 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell20);
                        float prevVal = Convert.ToSingle(prevData.Cell20);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 21, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 21, 0); }
                if (data.Cell21 != null) { SetXlsxCellValue(srcSheet, Range1, 22, data.Cell21.Value); }
                if (data.Cell22 != null) //摩尔比 大于0 小于2 
                {
                    if (data.Cell22.Value < 2)
                        SetXlsxCellValue(srcSheet, Range1, 23, data.Cell22.Value);
                }
                if (i == 12)
                {
                    if (data.Cell23 != null) { SetXlsxCellValue(srcSheet, Range1, 24, data.Cell23.Value); }//只记录最后一个值
                }
                if (data.Cell24 != null) { SetXlsxCellValue(srcSheet, Range1, 25, data.Cell24.Value); }

                if (data.Cell25 != null) { SetXlsxCellValue(srcSheet, Range1, 26, data.Cell25.Value); }
                if (data.Cell26 != null) { SetXlsxCellValue(srcSheet, Range1, 27, data.Cell26.Value); }
                if (data.Cell27 != null) { SetXlsxCellValue(srcSheet, Range1, 28, data.Cell27.Value); }
                if (data.Cell28 != null) { SetXlsxCellValue(srcSheet, Range1, 29, data.Cell28.Value); }
                //人工检测数据
                if (data.Cell29 != null) { SetXlsxCellValue(srcSheet, Range1, 30, data.Cell29.Value); }
                if (data.Cell30 != null) { SetXlsxCellValue(srcSheet, Range1, 31, data.Cell30.Value); }
                if (data.Cell31 != null) { SetXlsxCellValue(srcSheet, Range1, 32, data.Cell31.Value); }
                if (data.Cell32 != null) { SetXlsxCellValue(srcSheet, Range1, 33, data.Cell32.Value); }
                if (data.Cell33 != null) { SetXlsxCellValue(srcSheet, Range1, 34, data.Cell33.Value); }
                if (data.Cell34 != null) { SetXlsxCellValue(srcSheet, Range1, 35, data.Cell34.Value); }
                if (data.Cell35 != null) { SetXlsxCellValue(srcSheet, Range1, 36, data.Cell35.Value); }

                if (data.Cell36 != null) { SetXlsxCellValue(srcSheet, Range1, 37, data.Cell36.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell37 != null && prevData != null && prevData.Cell37 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell37);
                        float prevVal = Convert.ToSingle(prevData.Cell37);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 38, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 38, 0); }
                if (data.Cell38 != null) { SetXlsxCellValue(srcSheet, Range1, 39, data.Cell38.Value); }
                if (data.Cell39 != null) { SetXlsxCellValue(srcSheet, Range1, 40, data.Cell39.Value); }
                if (data.Cell40 != null) { SetXlsxCellValue(srcSheet, Range1, 41, data.Cell40.Value); }
                if (data.Cell41 != null) { SetXlsxCellValue(srcSheet, Range1, 42, data.Cell41.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell42 != null && prevData != null && prevData.Cell42 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell42);
                        float prevVal = Convert.ToSingle(prevData.Cell42);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 43, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 43, 0); }
                //if (data.Cell43 != null) { SetXlsxCellValue(srcSheet, Range1, 44, data.Cell43.Value); }
                //if (data.Cell44 != null) { SetXlsxCellValue(srcSheet, Range1, 45, data.Cell44.Value); }
                //if (data.Cell45 != null) { SetXlsxCellValue(srcSheet, Range1, 46, data.Cell45.Value); }
                //if (data.Cell46 != null) { SetXlsxCellValue(srcSheet, Range1, 47, data.Cell46.Value); }
                //if (data.Cell47 != null) { SetXlsxCellValue(srcSheet, Range1, 48, data.Cell47.Value); }
                //if (data.Cell48 != null) { SetXlsxCellValue(srcSheet, Range1, 49, data.Cell48.Value); }
                //if (data.Cell49 != null) { SetXlsxCellValue(srcSheet, Range1, 50, data.Cell49.Value); }
                //if (data.Cell50 != null) { SetXlsxCellValue(srcSheet, Range1, 51, data.Cell50.Value); }

                //Rang2
                if (data.Cell51 != null) { SetXlsxCellValue(srcSheet, Range2, 2, data.Cell51.Value); }
                if (data.Cell52 != null) { SetXlsxCellValue(srcSheet, Range2, 3, data.Cell52.Value); }
                if (data.Cell53 != null) { SetXlsxCellValue(srcSheet, Range2, 4, data.Cell53.Value); }
                if (data.Cell54 != null) { SetXlsxCellValue(srcSheet, Range2, 5, data.Cell54.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell55 != null && prevData != null && prevData.Cell55 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell55);
                        float prevVal = Convert.ToSingle(prevData.Cell55);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 6, result);
                    }
                }
                //人工检测数据
                if (data.Cell56 != null) { SetXlsxCellValue(srcSheet, Range2, 7, data.Cell56.Value); }
                if (data.Cell57 != null) { SetXlsxCellValue(srcSheet, Range2, 8, data.Cell57.Value); }
                if (data.Cell58 != null) { SetXlsxCellValue(srcSheet, Range2, 9, data.Cell58.Value); }
                if (data.Cell59 != null) { SetXlsxCellValue(srcSheet, Range2, 10, data.Cell59.Value); }
                if (data.Cell60 != null) { SetXlsxCellValue(srcSheet, Range2, 11, data.Cell60.Value); }

                if (data.Cell61 != null) { SetXlsxCellValue(srcSheet, Range2, 12, data.Cell61.Value); }
                if (data.Cell62 != null) { SetXlsxCellValue(srcSheet, Range2, 13, data.Cell62.Value); }
                if (data.Cell63 != null) { SetXlsxCellValue(srcSheet, Range2, 14, data.Cell63.Value); }
                if (data.Cell64 != null) { SetXlsxCellValue(srcSheet, Range2, 15, data.Cell64.Value); }
                if (data.Cell65 != null) { SetXlsxCellValue(srcSheet, Range2, 16, data.Cell65.Value); }
                if (data.Cell66 != null) { SetXlsxCellValue(srcSheet, Range2, 17, data.Cell66.Value); }
                if (data.Cell67 != null) { SetXlsxCellValue(srcSheet, Range2, 18, data.Cell67.Value); }
                if (data.Cell68 != null) { SetXlsxCellValue(srcSheet, Range2, 19, data.Cell68.Value); }
                if (data.Cell69 != null) { SetXlsxCellValue(srcSheet, Range2, 20, data.Cell69.Value); }
                if (data.Cell70 != null) { SetXlsxCellValue(srcSheet, Range2, 21, data.Cell70.Value); }
                if (data.Cell71 != null) { SetXlsxCellValue(srcSheet, Range2, 22, data.Cell71.Value); }
                if (data.Cell72 != null) { SetXlsxCellValue(srcSheet, Range2, 23, data.Cell72.Value); }
                if (data.Cell73 != null) { SetXlsxCellValue(srcSheet, Range2, 24, data.Cell73.Value); }
                if (data.Cell74 != null) { SetXlsxCellValue(srcSheet, Range2, 25, data.Cell74.Value); }
                if (data.Cell75 != null) { SetXlsxCellValue(srcSheet, Range2, 26, data.Cell75.Value); }
                if (data.Cell76 != null) { SetXlsxCellValue(srcSheet, Range2, 27, data.Cell76.Value); }
                if (data.Cell77 != null) { SetXlsxCellValue(srcSheet, Range2, 28, data.Cell77.Value); }
                if (data.Cell78 != null) { SetXlsxCellValue(srcSheet, Range2, 29, data.Cell78.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell79 != null && prevData != null && prevData.Cell79 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell79);
                        float prevVal = Convert.ToSingle(prevData.Cell79);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 30, 0); }
                if (data.Cell80 != null) { SetXlsxCellValue(srcSheet, Range2, 31, data.Cell80.Value); }
                if (data.Cell81 != null) { SetXlsxCellValue(srcSheet, Range2, 32, data.Cell81.Value); }
                //人工检测数据
                if (data.Cell82 != null) { SetXlsxCellValue(srcSheet, Range2, 33, data.Cell82.Value); }
                if (data.Cell83 != null) { SetXlsxCellValue(srcSheet, Range2, 34, data.Cell83.Value); }
                if (data.Cell84 != null) { SetXlsxCellValue(srcSheet, Range2, 35, data.Cell84.Value); }
                if (data.Cell85 != null) { SetXlsxCellValue(srcSheet, Range2, 36, data.Cell85.Value); }
                if (data.Cell86 != null) { SetXlsxCellValue(srcSheet, Range2, 37, data.Cell86.Value); }
                if (data.Cell87 != null) { SetXlsxCellValue(srcSheet, Range2, 38, data.Cell87.Value); }

                if (data.Cell88 != null) { SetXlsxCellValue(srcSheet, Range2, 39, data.Cell88.Value); }
                if (data.Cell89 != null) { SetXlsxCellValue(srcSheet, Range2, 40, data.Cell89.Value); }
                if (data.Cell90 != null) { SetXlsxCellValue(srcSheet, Range2, 41, data.Cell90.Value); }
                if (data.Cell91 != null) { SetXlsxCellValue(srcSheet, Range2, 42, data.Cell91.Value); }
                if (data.Cell92 != null) { SetXlsxCellValue(srcSheet, Range2, 43, data.Cell92.Value); }
                //if (data.Cell93 != null) { SetXlsxCellValue(srcSheet, Range2, 44, data.Cell93.Value); }
                //if (data.Cell94 != null) { SetXlsxCellValue(srcSheet, Range2, 45, data.Cell94.Value); }
                //if (data.Cell95 != null) { SetXlsxCellValue(srcSheet, Range2, 46, data.Cell95.Value); }
                //if (data.Cell96 != null) { SetXlsxCellValue(srcSheet, Range2, 47, data.Cell96.Value); }
                //if (data.Cell97 != null) { SetXlsxCellValue(srcSheet, Range2, 48, data.Cell97.Value); }
                //if (data.Cell98 != null) { SetXlsxCellValue(srcSheet, Range2, 49, data.Cell98.Value); }
                //if (data.Cell99 != null) { SetXlsxCellValue(srcSheet, Range2, 50, data.Cell99.Value); }
                //if (data.Cell100 != null) { SetXlsxCellValue(srcSheet, Range2, 51, data.Cell100.Value); }

                //Rang3
                if (data.Cell101 != null) { SetXlsxCellValue(srcSheet, Range3, 2, data.Cell101.Value * 1000); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell102 != null && prevData != null && prevData.Cell102 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell102);
                        float prevVal = Convert.ToSingle(prevData.Cell102);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 3, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 3, 0); }
                if (data.Cell103 != null) { SetXlsxCellValue(srcSheet, Range3, 4, data.Cell103.Value); }
                if (data.Cell104 != null) { SetXlsxCellValue(srcSheet, Range3, 5, data.Cell104.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell105 != null && prevData != null && prevData.Cell105 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell105);
                        float prevVal = Convert.ToSingle(prevData.Cell105);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 6, 0); }
                if (data.Cell106 != null) { SetXlsxCellValue(srcSheet, Range3, 7, data.Cell106.Value); }
                if (data.Cell107 != null) { SetXlsxCellValue(srcSheet, Range3, 8, data.Cell107.Value); }
                if (data.Cell108 != null) { SetXlsxCellValue(srcSheet, Range3, 9, data.Cell108.Value); }
                if (data.Cell109 != null) { SetXlsxCellValue(srcSheet, Range3, 10, data.Cell109.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell110 != null && prevData != null && prevData.Cell110 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell110);
                        float prevVal = Convert.ToSingle(prevData.Cell110);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 1, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 11, 0); }
                if (data.Cell111 != null) { SetXlsxCellValue(srcSheet, Range3, 12, data.Cell111.Value); }
                if (data.Cell112 != null) { SetXlsxCellValue(srcSheet, Range3, 13, data.Cell112.Value); }
                if (data.Cell113 != null) { SetXlsxCellValue(srcSheet, Range3, 14, data.Cell113.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell114 != null && prevData != null && prevData.Cell114 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell114);
                        float prevVal = Convert.ToSingle(prevData.Cell114);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 15, 0); }
                if (data.Cell115 != null) { SetXlsxCellValue(srcSheet, Range3, 16, data.Cell115.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell116 != null && prevData != null && prevData.Cell116 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell116);
                        float prevVal = Convert.ToSingle(prevData.Cell116);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 17, 0); }
                if (data.Cell117 != null) { SetXlsxCellValue(srcSheet, Range3, 18, data.Cell117.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell118 != null && prevData != null && prevData.Cell118 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell118);
                        float prevVal = Convert.ToSingle(prevData.Cell118);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 19, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 19, 0); }
                if (data.Cell119 != null) { SetXlsxCellValue(srcSheet, Range3, 20, data.Cell119.Value); }
                if (data.Cell120 != null) { SetXlsxCellValue(srcSheet, Range3, 21, data.Cell120.Value); }
                if (data.Cell121 != null) { SetXlsxCellValue(srcSheet, Range3, 22, data.Cell121.Value); }
                if (data.Cell122 != null) { SetXlsxCellValue(srcSheet, Range3, 23, data.Cell122.Value); }
                if (data.Cell123 != null) { SetXlsxCellValue(srcSheet, Range3, 24, data.Cell123.Value); }
                if (data.Cell124 != null) { SetXlsxCellValue(srcSheet, Range3, 25, data.Cell124.Value); }
                if (data.Cell125 != null) { SetXlsxCellValue(srcSheet, Range3, 26, data.Cell125.Value); }
                if (data.Cell126 != null) { SetXlsxCellValue(srcSheet, Range3, 27, data.Cell126.Value); }
                if (data.Cell127 != null) { SetXlsxCellValue(srcSheet, Range3, 28, data.Cell127.Value); }
                if (data.Cell128 != null) { SetXlsxCellValue(srcSheet, Range3, 29, data.Cell128.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell129 != null && prevData != null && prevData.Cell129 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell129);
                        float prevVal = Convert.ToSingle(prevData.Cell129);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 30, 0); }
                if (data.Cell130 != null) { SetXlsxCellValue(srcSheet, Range3, 31, data.Cell130.Value); }
                if (data.Cell131 != null) { SetXlsxCellValue(srcSheet, Range3, 32, data.Cell131.Value); }
                if (i != 0)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell132 != null && prevData != null && prevData.Cell132 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell132);
                        float prevVal = Convert.ToSingle(prevData.Cell132);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 33, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 33, 0); }
                if (data.Cell133 != null) { SetXlsxCellValue(srcSheet, Range3, 34, data.Cell133.Value); }
                if (data.Cell134 != null) { SetXlsxCellValue(srcSheet, Range3, 35, data.Cell134.Value); }
                //人工检测数据
                if (data.Cell135 != null) { SetXlsxCellValue(srcSheet, Range3, 36, data.Cell135.Value); }
                if (data.Cell136 != null) { SetXlsxCellValue(srcSheet, Range3, 37, data.Cell136.Value); }
                if (data.Cell137 != null) { SetXlsxCellValue(srcSheet, Range3, 38, data.Cell137.Value); }
                if (data.Cell138 != null) { SetXlsxCellValue(srcSheet, Range3, 39, data.Cell138.Value); }
                if (data.Cell139 != null) { SetXlsxCellValue(srcSheet, Range3, 40, data.Cell139.Value); }
                if (data.Cell140 != null) { SetXlsxCellValue(srcSheet, Range3, 41, data.Cell140.Value); }
                if (data.Cell141 != null) { SetXlsxCellValue(srcSheet, Range3, 42, data.Cell141.Value); }

                //if (data.Cell142 != null) { SetXlsxCellValue(srcSheet, Range3, 43, data.Cell142.Value); }
                //if (data.Cell143 != null) { SetXlsxCellValue(srcSheet, Range3, 44, data.Cell143.Value); }
                //if (data.Cell144 != null) { SetXlsxCellValue(srcSheet, Range3, 45, data.Cell144.Value); }
                //if (data.Cell145 != null) { SetXlsxCellValue(srcSheet, Range3, 46, data.Cell145.Value); }
                //if (data.Cell146 != null) { SetXlsxCellValue(srcSheet, Range3, 47, data.Cell146.Value); }
                //if (data.Cell147 != null) { SetXlsxCellValue(srcSheet, Range3, 48, data.Cell147.Value); }
                //if (data.Cell148 != null) { SetXlsxCellValue(srcSheet, Range3, 49, data.Cell148.Value); }
                //if (data.Cell149 != null) { SetXlsxCellValue(srcSheet, Range3, 50, data.Cell149.Value); }
                //if (data.Cell150 != null) { SetXlsxCellValue(srcSheet, Range3, 51, data.Cell150.Value); }

            }
            return true;
        }

        /// <summary>
        /// 写Xlsx数据  夜班
        /// </summary>
        private static bool NightWriteExcel(XSSFWorkbook srcWorkbook, SourceData?[] dataList, DateTime ReportDataTime)
        {

            ISheet srcSheet = srcWorkbook.GetSheetAt(2); //实际要写的表
            string Temp = ReportDataTime.Date.ToString("yyyy-MM-dd");
            SetXlsxCellString(srcSheet, 51, 1, Temp);//记录日期
            srcSheet.ForceFormulaRecalculation = false;//批量写入关闭公式自动计算，大幅提升写入速度
            for (int i = 12; i < 25; i++)
            {
                var data = dataList.ElementAt(i);
                if (data == null) continue; // 如果 data 为空则跳过

                int Range1 = 5 + i - 12;
                int Range2 = 21 + i - 12;
                int Range3 = 38 + i - 12;

                // 从Excel第2列开始写入
                //Rang1 
                if (data.Cell1 != null) { SetXlsxCellValue(srcSheet, Range1, 2, data.Cell1.Value); }
                if (data.Cell2 != null) { SetXlsxCellValue(srcSheet, Range1, 3, data.Cell2.Value); }
                if (data.Cell3 != null) { SetXlsxCellValue(srcSheet, Range1, 4, data.Cell3.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell4 != null && prevData != null && prevData.Cell4 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell4);
                        float prevVal = Convert.ToSingle(prevData.Cell4);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 5, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 5, 0); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell5 != null && prevData != null && prevData.Cell5 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell5);
                        float prevVal = Convert.ToSingle(prevData.Cell5);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 6, 0); }
                if (data.Cell6 != null) { SetXlsxCellValue(srcSheet, Range1, 7, data.Cell6.Value); }
                if (data.Cell7 != null) { SetXlsxCellValue(srcSheet, Range1, 8, data.Cell7.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell8 != null && prevData != null && prevData.Cell8 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell8);
                        float prevVal = Convert.ToSingle(prevData.Cell8);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 9, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 9, 0); }
                if (data.Cell9 != null) { SetXlsxCellValue(srcSheet, Range1, 10, data.Cell9.Value); }
                if (data.Cell10 != null) { SetXlsxCellValue(srcSheet, Range1, 11, data.Cell10.Value); }
                if (data.Cell11 != null) { SetXlsxCellValue(srcSheet, Range1, 12, data.Cell11.Value); }
                if (data.Cell12 != null) { SetXlsxCellValue(srcSheet, Range1, 13, data.Cell12.Value); }
                if (data.Cell13 != null) { SetXlsxCellValue(srcSheet, Range1, 14, data.Cell13.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell14 != null && prevData != null && prevData.Cell14 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell14);
                        float prevVal = Convert.ToSingle(prevData.Cell14);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 15, 0); }
                if (data.Cell15 != null) { SetXlsxCellValue(srcSheet, Range1, 16, data.Cell15.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell16 != null && prevData != null && prevData.Cell16 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell16);
                        float prevVal = Convert.ToSingle(prevData.Cell16);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 17, 0); }
                if (data.Cell17 != null) { SetXlsxCellValue(srcSheet, Range1, 18, data.Cell17.Value); }
                if (data.Cell18 != null) { SetXlsxCellValue(srcSheet, Range1, 19, data.Cell18.Value); }
                if (data.Cell19 != null) { SetXlsxCellValue(srcSheet, Range1, 20, data.Cell19.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell20 != null && prevData != null && prevData.Cell20 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell20);
                        float prevVal = Convert.ToSingle(prevData.Cell20);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 21, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 21, 0); }
                if (data.Cell21 != null) { SetXlsxCellValue(srcSheet, Range1, 22, data.Cell21.Value); }
                if (data.Cell22 != null) //摩尔比 大于0 小于2 
                {
                    if (data.Cell22.Value < 2)
                        SetXlsxCellValue(srcSheet, Range1, 23, data.Cell22.Value);
                }
                if (i == 24)
                {
                    if (data.Cell23 != null) { SetXlsxCellValue(srcSheet, Range1, 24, data.Cell23.Value); }//只记录最后一个值
                }
                if (data.Cell24 != null) { SetXlsxCellValue(srcSheet, Range1, 25, data.Cell24.Value); }

                if (data.Cell25 != null) { SetXlsxCellValue(srcSheet, Range1, 26, data.Cell25.Value); }
                if (data.Cell26 != null) { SetXlsxCellValue(srcSheet, Range1, 27, data.Cell26.Value); }
                if (data.Cell27 != null) { SetXlsxCellValue(srcSheet, Range1, 28, data.Cell27.Value); }
                if (data.Cell28 != null) { SetXlsxCellValue(srcSheet, Range1, 29, data.Cell28.Value); }
                //人工检测数据
                if (data.Cell29 != null) { SetXlsxCellValue(srcSheet, Range1, 30, data.Cell29.Value); }
                if (data.Cell30 != null) { SetXlsxCellValue(srcSheet, Range1, 31, data.Cell30.Value); }
                if (data.Cell31 != null) { SetXlsxCellValue(srcSheet, Range1, 32, data.Cell31.Value); }
                if (data.Cell32 != null) { SetXlsxCellValue(srcSheet, Range1, 33, data.Cell32.Value); }
                if (data.Cell33 != null) { SetXlsxCellValue(srcSheet, Range1, 34, data.Cell33.Value); }
                if (data.Cell34 != null) { SetXlsxCellValue(srcSheet, Range1, 35, data.Cell34.Value); }
                if (data.Cell35 != null) { SetXlsxCellValue(srcSheet, Range1, 36, data.Cell35.Value); }

                if (data.Cell36 != null) { SetXlsxCellValue(srcSheet, Range1, 37, data.Cell36.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell37 != null && prevData != null && prevData.Cell37 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell37);
                        float prevVal = Convert.ToSingle(prevData.Cell37);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 38, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 38, 0); }
                if (data.Cell38 != null) { SetXlsxCellValue(srcSheet, Range1, 39, data.Cell38.Value); }
                if (data.Cell39 != null) { SetXlsxCellValue(srcSheet, Range1, 40, data.Cell39.Value); }
                if (data.Cell40 != null) { SetXlsxCellValue(srcSheet, Range1, 41, data.Cell40.Value); }
                if (data.Cell41 != null) { SetXlsxCellValue(srcSheet, Range1, 42, data.Cell41.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell42 != null && prevData != null && prevData.Cell42 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell42);
                        float prevVal = Convert.ToSingle(prevData.Cell42);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range1, 43, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range1, 43, 0); }
                //if (data.Cell43 != null) { SetXlsxCellValue(srcSheet, Range1, 44, data.Cell43.Value); }
                //if (data.Cell44 != null) { SetXlsxCellValue(srcSheet, Range1, 45, data.Cell44.Value); }
                //if (data.Cell45 != null) { SetXlsxCellValue(srcSheet, Range1, 46, data.Cell45.Value); }
                //if (data.Cell46 != null) { SetXlsxCellValue(srcSheet, Range1, 47, data.Cell46.Value); }
                //if (data.Cell47 != null) { SetXlsxCellValue(srcSheet, Range1, 48, data.Cell47.Value); }
                //if (data.Cell48 != null) { SetXlsxCellValue(srcSheet, Range1, 49, data.Cell48.Value); }
                //if (data.Cell49 != null) { SetXlsxCellValue(srcSheet, Range1, 50, data.Cell49.Value); }
                //if (data.Cell50 != null) { SetXlsxCellValue(srcSheet, Range1, 51, data.Cell50.Value); }

                //Rang2
                if (data.Cell51 != null) { SetXlsxCellValue(srcSheet, Range2, 2, data.Cell51.Value); }
                if (data.Cell52 != null) { SetXlsxCellValue(srcSheet, Range2, 3, data.Cell52.Value); }
                if (data.Cell53 != null) { SetXlsxCellValue(srcSheet, Range2, 4, data.Cell53.Value); }
                if (data.Cell54 != null) { SetXlsxCellValue(srcSheet, Range2, 5, data.Cell54.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell55 != null && prevData != null && prevData.Cell55 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell55);
                        float prevVal = Convert.ToSingle(prevData.Cell55);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 6, 0); }
                //人工检测数据
                if (data.Cell56 != null) { SetXlsxCellValue(srcSheet, Range2, 7, data.Cell56.Value); }
                if (data.Cell57 != null) { SetXlsxCellValue(srcSheet, Range2, 8, data.Cell57.Value); }
                if (data.Cell58 != null) { SetXlsxCellValue(srcSheet, Range2, 9, data.Cell58.Value); }
                if (data.Cell59 != null) { SetXlsxCellValue(srcSheet, Range2, 10, data.Cell59.Value); }
                if (data.Cell60 != null) { SetXlsxCellValue(srcSheet, Range2, 11, data.Cell60.Value); }

                if (data.Cell61 != null) { SetXlsxCellValue(srcSheet, Range2, 12, data.Cell61.Value); }
                if (data.Cell62 != null) { SetXlsxCellValue(srcSheet, Range2, 13, data.Cell62.Value); }
                if (data.Cell63 != null) { SetXlsxCellValue(srcSheet, Range2, 14, data.Cell63.Value); }
                if (data.Cell64 != null) { SetXlsxCellValue(srcSheet, Range2, 15, data.Cell64.Value); }
                if (data.Cell65 != null) { SetXlsxCellValue(srcSheet, Range2, 16, data.Cell65.Value); }
                if (data.Cell66 != null) { SetXlsxCellValue(srcSheet, Range2, 17, data.Cell66.Value); }
                if (data.Cell67 != null) { SetXlsxCellValue(srcSheet, Range2, 18, data.Cell67.Value); }
                if (data.Cell68 != null) { SetXlsxCellValue(srcSheet, Range2, 19, data.Cell68.Value); }
                if (data.Cell69 != null) { SetXlsxCellValue(srcSheet, Range2, 20, data.Cell69.Value); }
                if (data.Cell70 != null) { SetXlsxCellValue(srcSheet, Range2, 21, data.Cell70.Value); }
                if (data.Cell71 != null) { SetXlsxCellValue(srcSheet, Range2, 22, data.Cell71.Value); }
                if (data.Cell72 != null) { SetXlsxCellValue(srcSheet, Range2, 23, data.Cell72.Value); }
                if (data.Cell73 != null) { SetXlsxCellValue(srcSheet, Range2, 24, data.Cell73.Value); }
                if (data.Cell74 != null) { SetXlsxCellValue(srcSheet, Range2, 25, data.Cell74.Value); }
                if (data.Cell75 != null) { SetXlsxCellValue(srcSheet, Range2, 26, data.Cell75.Value); }
                if (data.Cell76 != null) { SetXlsxCellValue(srcSheet, Range2, 27, data.Cell76.Value); }
                if (data.Cell77 != null) { SetXlsxCellValue(srcSheet, Range2, 28, data.Cell77.Value); }
                if (data.Cell78 != null) { SetXlsxCellValue(srcSheet, Range2, 29, data.Cell78.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell79 != null && prevData != null && prevData.Cell79 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell79);
                        float prevVal = Convert.ToSingle(prevData.Cell79);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range2, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range2, 30, 0); }
                if (data.Cell80 != null) { SetXlsxCellValue(srcSheet, Range2, 31, data.Cell80.Value); }
                if (data.Cell81 != null) { SetXlsxCellValue(srcSheet, Range2, 32, data.Cell81.Value); }
                //人工检测数据
                if (data.Cell82 != null) { SetXlsxCellValue(srcSheet, Range2, 33, data.Cell82.Value); }
                if (data.Cell83 != null) { SetXlsxCellValue(srcSheet, Range2, 34, data.Cell83.Value); }
                if (data.Cell84 != null) { SetXlsxCellValue(srcSheet, Range2, 35, data.Cell84.Value); }
                if (data.Cell85 != null) { SetXlsxCellValue(srcSheet, Range2, 36, data.Cell85.Value); }
                if (data.Cell86 != null) { SetXlsxCellValue(srcSheet, Range2, 37, data.Cell86.Value); }
                if (data.Cell87 != null) { SetXlsxCellValue(srcSheet, Range2, 38, data.Cell87.Value); }

                if (data.Cell88 != null) { SetXlsxCellValue(srcSheet, Range2, 39, data.Cell88.Value); }
                if (data.Cell89 != null) { SetXlsxCellValue(srcSheet, Range2, 40, data.Cell89.Value); }
                if (data.Cell90 != null) { SetXlsxCellValue(srcSheet, Range2, 41, data.Cell90.Value); }
                if (data.Cell91 != null) { SetXlsxCellValue(srcSheet, Range2, 42, data.Cell91.Value); }
                if (data.Cell92 != null) { SetXlsxCellValue(srcSheet, Range2, 43, data.Cell92.Value); }
                //if (data.Cell93 != null) { SetXlsxCellValue(srcSheet, Range2, 44, data.Cell93.Value); }
                //if (data.Cell94 != null) { SetXlsxCellValue(srcSheet, Range2, 45, data.Cell94.Value); }
                //if (data.Cell95 != null) { SetXlsxCellValue(srcSheet, Range2, 46, data.Cell95.Value); }
                //if (data.Cell96 != null) { SetXlsxCellValue(srcSheet, Range2, 47, data.Cell96.Value); }
                //if (data.Cell97 != null) { SetXlsxCellValue(srcSheet, Range2, 48, data.Cell97.Value); }
                //if (data.Cell98 != null) { SetXlsxCellValue(srcSheet, Range2, 49, data.Cell98.Value); }
                //if (data.Cell99 != null) { SetXlsxCellValue(srcSheet, Range2, 50, data.Cell99.Value); }
                //if (data.Cell100 != null) { SetXlsxCellValue(srcSheet, Range2, 51, data.Cell100.Value); }

                //Rang3
                if (data.Cell101 != null) { SetXlsxCellValue(srcSheet, Range3, 2, data.Cell101.Value * 1000); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell102 != null && prevData != null && prevData.Cell102 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell102);
                        float prevVal = Convert.ToSingle(prevData.Cell102);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 3, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 3, 0); }
                if (data.Cell103 != null) { SetXlsxCellValue(srcSheet, Range3, 4, data.Cell103.Value); }
                if (data.Cell104 != null) { SetXlsxCellValue(srcSheet, Range3, 5, data.Cell104.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell105 != null && prevData != null && prevData.Cell105 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell105);
                        float prevVal = Convert.ToSingle(prevData.Cell105);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 6, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 6, 0); }
                if (data.Cell106 != null) { SetXlsxCellValue(srcSheet, Range3, 7, data.Cell106.Value); }
                if (data.Cell107 != null) { SetXlsxCellValue(srcSheet, Range3, 8, data.Cell107.Value); }
                if (data.Cell108 != null) { SetXlsxCellValue(srcSheet, Range3, 9, data.Cell108.Value); }
                if (data.Cell109 != null) { SetXlsxCellValue(srcSheet, Range3, 10, data.Cell109.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell110 != null && prevData != null && prevData.Cell110 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell110);
                        float prevVal = Convert.ToSingle(prevData.Cell110);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 1, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 11, 0); }
                if (data.Cell111 != null) { SetXlsxCellValue(srcSheet, Range3, 12, data.Cell111.Value); }
                if (data.Cell112 != null) { SetXlsxCellValue(srcSheet, Range3, 13, data.Cell112.Value); }
                if (data.Cell113 != null) { SetXlsxCellValue(srcSheet, Range3, 14, data.Cell113.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell114 != null && prevData != null && prevData.Cell114 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell114);
                        float prevVal = Convert.ToSingle(prevData.Cell114);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 15, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 15, 0); }
                if (data.Cell115 != null) { SetXlsxCellValue(srcSheet, Range3, 16, data.Cell115.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell116 != null && prevData != null && prevData.Cell116 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell116);
                        float prevVal = Convert.ToSingle(prevData.Cell116);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 17, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 17, 0); }
                if (data.Cell117 != null) { SetXlsxCellValue(srcSheet, Range3, 18, data.Cell117.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell118 != null && prevData != null && prevData.Cell118 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell118);
                        float prevVal = Convert.ToSingle(prevData.Cell118);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 19, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 19, 0); }
                if (data.Cell119 != null) { SetXlsxCellValue(srcSheet, Range3, 20, data.Cell119.Value); }
                if (data.Cell120 != null) { SetXlsxCellValue(srcSheet, Range3, 21, data.Cell120.Value); }
                if (data.Cell121 != null) { SetXlsxCellValue(srcSheet, Range3, 22, data.Cell121.Value); }
                if (data.Cell122 != null) { SetXlsxCellValue(srcSheet, Range3, 23, data.Cell122.Value); }
                if (data.Cell123 != null) { SetXlsxCellValue(srcSheet, Range3, 24, data.Cell123.Value); }
                if (data.Cell124 != null) { SetXlsxCellValue(srcSheet, Range3, 25, data.Cell124.Value); }
                if (data.Cell125 != null) { SetXlsxCellValue(srcSheet, Range3, 26, data.Cell125.Value); }
                if (data.Cell126 != null) { SetXlsxCellValue(srcSheet, Range3, 27, data.Cell126.Value); }
                if (data.Cell127 != null) { SetXlsxCellValue(srcSheet, Range3, 28, data.Cell127.Value); }
                if (data.Cell128 != null) { SetXlsxCellValue(srcSheet, Range3, 29, data.Cell128.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell129 != null && prevData != null && prevData.Cell129 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell129);
                        float prevVal = Convert.ToSingle(prevData.Cell129);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 30, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 30, 0); }
                if (data.Cell130 != null) { SetXlsxCellValue(srcSheet, Range3, 31, data.Cell130.Value); }
                if (data.Cell131 != null) { SetXlsxCellValue(srcSheet, Range3, 32, data.Cell131.Value); }
                if (i != 12)// 每小时的差值
                {
                    var prevData = dataList.ElementAt(i - 1);
                    if (data.Cell132 != null && prevData != null && prevData.Cell132 != null)
                    {
                        float currentVal = Convert.ToSingle(data.Cell132);
                        float prevVal = Convert.ToSingle(prevData.Cell132);
                        float result = (float)Math.Round((currentVal - prevVal) / 1000, 2);
                        SetXlsxCellValue(srcSheet, Range3, 33, result);
                    }
                }
                else { SetXlsxCellValue(srcSheet, Range3, 33, 0); }
                if (data.Cell133 != null) { SetXlsxCellValue(srcSheet, Range3, 34, data.Cell133.Value); }
                if (data.Cell134 != null) { SetXlsxCellValue(srcSheet, Range3, 35, data.Cell134.Value); }
                //人工检测数据
                if (data.Cell135 != null) { SetXlsxCellValue(srcSheet, Range3, 36, data.Cell135.Value); }
                if (data.Cell136 != null) { SetXlsxCellValue(srcSheet, Range3, 37, data.Cell136.Value); }
                if (data.Cell137 != null) { SetXlsxCellValue(srcSheet, Range3, 38, data.Cell137.Value); }
                if (data.Cell138 != null) { SetXlsxCellValue(srcSheet, Range3, 39, data.Cell138.Value); }
                if (data.Cell139 != null) { SetXlsxCellValue(srcSheet, Range3, 40, data.Cell139.Value); }
                if (data.Cell140 != null) { SetXlsxCellValue(srcSheet, Range3, 41, data.Cell140.Value); }
                if (data.Cell141 != null) { SetXlsxCellValue(srcSheet, Range3, 42, data.Cell141.Value); }

                //if (data.Cell142 != null) { SetXlsxCellValue(srcSheet, Range3, 43, data.Cell142.Value); }
                //if (data.Cell143 != null) { SetXlsxCellValue(srcSheet, Range3, 44, data.Cell143.Value); }
                //if (data.Cell144 != null) { SetXlsxCellValue(srcSheet, Range3, 45, data.Cell144.Value); }
                //if (data.Cell145 != null) { SetXlsxCellValue(srcSheet, Range3, 46, data.Cell145.Value); }
                //if (data.Cell146 != null) { SetXlsxCellValue(srcSheet, Range3, 47, data.Cell146.Value); }
                //if (data.Cell147 != null) { SetXlsxCellValue(srcSheet, Range3, 48, data.Cell147.Value); }
                //if (data.Cell148 != null) { SetXlsxCellValue(srcSheet, Range3, 49, data.Cell148.Value); }
                //if (data.Cell149 != null) { SetXlsxCellValue(srcSheet, Range3, 50, data.Cell149.Value); }
                //if (data.Cell150 != null) { SetXlsxCellValue(srcSheet, Range3, 51, data.Cell150.Value); }

            }
            return true;
        }


        private static bool MonthWriteExcel(XSSFWorkbook srcWorkbook, DayWorkBook monthWorkBookData, string reportedTime)
        {
            return true;
        }

        private static bool YearWriteExcel(XSSFWorkbook srcWorkbook, DayWorkBook yearWorkBookData, string reportedTime)
        {
            return true;
        }
        private static bool WeekWriteExcel(XSSFWorkbook srcWorkbook, WeekWorkBook weekWorkBookData, string reportedTime)
        {
            return true;
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
