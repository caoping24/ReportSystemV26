using CenterBackend.IFileService;
using CenterBackend.Logging;
using CenterBackend.Models;
using ICSharpCode.SharpZipLib.Zip;
using System.Web;
namespace CenterBackend.Services
{
    /// <summary>
    /// 文件操作服务实现类 - 实现所有文件相关的具体逻辑
    /// </summary>
    public class FileService : IFileServices
    {
        private readonly IAppLogger _logger;
        public FileService(IAppLogger _IAppLogger)
        {
            this._logger = _IAppLogger;
        }
        /// <summary>
        /// 创建文件夹兼容单层/多层，路径存在则跳过，无异常抛出
        /// </summary>
        /// <param name="folderPath">待创建的文件夹完整路径</param>
        public void CreateFolder(string folderPath)
        {
            if (!string.IsNullOrWhiteSpace(folderPath) && !Directory.Exists(folderPath))//判空+判存在
            {
                Directory.CreateDirectory(folderPath);
            }
        }

        /// <summary>
        /// 复制文件核心封装自动创建目标文件夹
        /// </summary>
        /// <param name="sourceFilePath">源文件完整路径</param>
        /// <param name="targetFilePath">目标文件完整路径</param>
        /// <param name="overwrite">是否覆盖已存在的目标文件，默认:true</param>
        /// <returns>复制成功返回true，失败返回false</returns>
        public bool CopyFile(string sourceFilePath, string targetFilePath, bool overwrite = true)
        {
            try
            {
                if (!File.Exists(sourceFilePath)) throw new FileNotFoundException($"源文件不存在：{sourceFilePath}");//校验源文件

                string targetFolder = Path.GetDirectoryName(targetFilePath);//提取目标文件夹路径
                CreateFolder(targetFolder);//自动创建目标文件夹

                File.Copy(sourceFilePath, targetFilePath, overwrite);
                return true;
            }
            catch
            {
                return false;//异常则返回失败
            }
        }

        /// <summary>
        /// 压缩指定文件夹内所有内容为Zip包含子文件夹+保持原结构
        /// </summary>
        /// <param name="sourceFolderPath">待压缩的源文件夹路径</param>
        /// <param name="zipSavePath">压缩包保存的完整路径(含文件名.zip)</param>
        /// <returns>压缩成功返回true，失败返回false</returns>
        public bool CompressFolderToZip(string sourceFolderPath, string zipSavePath)
        {
            try
            {
                if (!Directory.Exists(sourceFolderPath)) throw new DirectoryNotFoundException($"文件夹不存在：{sourceFolderPath}");//校验源文件夹

                string zipFolder = Path.GetDirectoryName(zipSavePath);//提取压缩包所在文件夹
                CreateFolder(zipFolder);//自动创建存储目录

                using (var fs = new FileStream(zipSavePath, FileMode.Create, FileAccess.Write))
                using (var zipStream = new ZipOutputStream(fs) { IsStreamOwner = true })
                {
                    zipStream.SetLevel(6);//0-9，6=速度+压缩率最优平衡
                    CompressDirectory(sourceFolderPath, zipStream, sourceFolderPath);//递归压缩文件/子文件夹
                    zipStream.Finish();
                }
                return File.Exists(zipSavePath);//校验压缩包是否生成
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 单个文件下载,压缩为独立Zip包
        /// </summary>
        /// <param name="sourceFilePath">待压缩的源文件完整路径</param>
        /// <param name="zipSavePath">压缩包保存的完整路径(含文件名.zip)</param>
        /// <returns>压缩成功返回true，失败返回false</returns>
        public bool CompressSingleFileToZip(string sourceFilePath, string zipSavePath)
        {
            try
            {
                if (!File.Exists(sourceFilePath)) throw new FileNotFoundException($"文件不存在：{sourceFilePath}");//校验源文件

                string zipFolder = Path.GetDirectoryName(zipSavePath);//提取压缩包所在文件夹
                CreateFolder(zipFolder);//自动创建存储目录

                // 流式写入，缓冲区读取，适配大文件
                using (var fs = new FileStream(zipSavePath, FileMode.Create, FileAccess.Write))
                using (var zipStream = new ZipOutputStream(fs) { IsStreamOwner = true })
                {
                    zipStream.SetLevel(6);//最优压缩级别
                    var fileInfo = new FileInfo(sourceFilePath);
                    var zipEntry = new ZipEntry(fileInfo.Name) { DateTime = DateTime.Now };//创建压缩项
                    zipStream.PutNextEntry(zipEntry);

                    byte[] buffer = new byte[4096];//4K缓冲区，平衡性能与内存
                    using (var fileStream = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
                    {
                        int bytesRead;
                        while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            zipStream.Write(buffer, 0, bytesRead);
                        }
                    }
                    zipStream.CloseEntry();
                    zipStream.Finish();
                }
                return File.Exists(zipSavePath);//校验压缩包是否生成
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 单个文件直接下载,不压缩
        /// </summary>
        /// <param name="fileName">文件名称（如：test.xlsx、数据.csv）</param>
        /// <param name="sourceFilePath">文件所在目录</param>
        /// <returns>返回文件流+文件名，供控制器直接返回，null=文件不存在</returns>
        public (FileStream? stream, string? encodeFileName) DownloadSingleFile(string sourceFilePath, string fileName)
        {
            try
            {
                string filePath = Path.Combine(sourceFilePath, fileName);
                if (!File.Exists(filePath)) return (null, null);
                var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);// 只读方式打开文件流，共享读取（文件被占用也能下载），流式返回，不加载到内存
                string encodeFileName = HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8);// 解决中文文件名下载乱码问题
                return (fileStream, encodeFileName);
            }
            catch
            {
                return (null, null);
            }
        }

        #region 私有辅助方法
        /// <summary>
        /// 递归压缩文件夹内所有文件/子文件夹，保持原目录相对结构
        /// </summary>
        /// <param name="folderPath">当前遍历的文件夹路径</param>
        /// <param name="zipStream">压缩输出流</param>
        /// <param name="rootFolder">压缩根目录，用于生成相对路径</param>
        private void CompressDirectory(string folderPath, ZipOutputStream zipStream, string rootFolder)
        {
            //遍历当前文件夹所有文件
            foreach (string file in Directory.GetFiles(folderPath))
            {
                var fileInfo = new FileInfo(file);
                string entryName = file.Substring(rootFolder.Length + 1);//生成压缩包内相对路径
                var zipEntry = new ZipEntry(entryName) { DateTime = DateTime.Now };
                zipStream.PutNextEntry(zipEntry);

                byte[] buffer = new byte[4096];//4K缓冲区
                using (var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    int bytesRead;
                    while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        zipStream.Write(buffer, 0, bytesRead);
                    }
                }
                zipStream.CloseEntry();
            }

            //递归遍历子文件夹
            foreach (string dir in Directory.GetDirectories(folderPath))
            {
                CompressDirectory(dir, zipStream, rootFolder);
            }
        }
        #endregion
        /// <summary>
        /// 获取指定日期的文件路径+文件名（日报/周报/年报）
        /// 日报按目标日期年月 | 周报按本周一的年月+年周序号 | 年报按年份
        /// </summary>
        public FilePathAndName GetDateFolderPathAndName(string rootPath, DateTime targetDate)
        {
            var result = new FilePathAndName();
            if (string.IsNullOrWhiteSpace(rootPath)) return result;

            DateTime currentDate = targetDate.Date;//去时分秒，纯日期计算
            DateTime weekFirstDay = GetWeekFirstDay(currentDate);//获取该日期的本周一
            int weekBelongYear = weekFirstDay.Year;//本周一 归属的年份
            int weekBelongMonth = weekFirstDay.Month;//本周一 归属的月份
            int weekNumberInYear = GetWeekNumberInYear(weekFirstDay);//该周在当年的周序号

            // ===== 日报表
            result.DailyFilesPath = Path.Combine(rootPath, "日报表", $"{currentDate.Year}-{currentDate.Month:00}");
            result.DailyFileName = $"日报表-{currentDate.Year}-{currentDate.Month:00}-{currentDate.Day:00}.xlsx";
            result.DailyFilesFullPath = Path.Combine(result.DailyFilesPath, result.DailyFileName);
            // ===== 月报表
            result.MonthlyFilesPath = Path.Combine(rootPath, "月报表", $"{currentDate.Year}");
            result.MonthlyFileName = $"月报表-{currentDate.Year}-{currentDate.Month:00}.xlsx";
            result.MonthlyFilesFullPath = Path.Combine(result.MonthlyFilesPath, result.MonthlyFileName);

            // ===== 周报表 路径+文件名 (核心：按[本周一]的年月归类 =====
            result.WeeklyFilesPath = Path.Combine(rootPath, "周报表", $"{weekBelongYear}-{weekBelongMonth:00}");
            result.WeeklyFileName = $"周报表-{weekBelongYear}年{weekNumberInYear:00}周.xlsx";
            result.WeeklyFilesFullPath = Path.Combine(result.WeeklyFilesPath, result.WeeklyFileName);

            // ===== 年报表
            result.YearlyFilesPath = Path.Combine(rootPath, "年报表");
            result.YearlyFileName = $"年报表-{currentDate.Year}.xlsx";
            result.YearlyFilesFullPath = Path.Combine(result.YearlyFilesPath, result.YearlyFileName);

            return result;
        }


        /// <summary>
        /// 按日期创建年/月/周三级文件夹(存在则不创建)
        /// </summary>
        /// <param name="rootPath">存储根路径</param>
        /// <param name="targetDate">目标日期</param>
        public void CreateDateFolder(string rootPath, DateTime targetDate)
        {
            var result = GetDateFolderPathAndName(rootPath, targetDate);
            if (string.IsNullOrWhiteSpace(result.DailyFilesPath)) return;
            CreateFolder(result.DailyFilesPath);
            CreateFolder(result.MonthlyFilesPath);
            CreateFolder(result.WeeklyFilesPath);
            CreateFolder(result.YearlyFilesPath);
        }
        #region 私有辅助方法
        /// <summary>
        /// 获取指定日期的本周第一天(周一为周首)
        /// </summary>
        /// <param name="dt">目标日期</param>
        /// <returns>本周一日期</returns>
        private DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = dt.DayOfWeek - DayOfWeek.Monday;//计算与周一的差值
            if (diff < 0) diff += 7;//周日处理为-1，补7天
            return dt.AddDays(-diff).Date;
        }
        /// <summary>
        /// 根据目标日期，先获取该周第一天(周一)，再计算【该周是当年的第几周】
        /// </summary>
        /// <param name="dt">任意目标日期</param>
        /// <returns>该周在当年的周序号 1~53</returns>
        private int GetWeekNumberInYear(DateTime dt)
        {

            DateTime weekFirstDay = GetWeekFirstDay(dt);

            DateTime yearFirstDay = new DateTime(weekFirstDay.Year, 1, 1);
            // 当年的第一个周一= 当年元旦的本周一
            DateTime yearFirstWeekDay = GetWeekFirstDay(yearFirstDay);
            // 两个周一的间隔天数 / 7  = 周差，+1得到周序号（从1开始）
            int daysDiff = (int)(weekFirstDay - yearFirstWeekDay).TotalDays;
            int weekNumber = (daysDiff / 7) + 1;
            return weekNumber;
        }
        #endregion
    }
}