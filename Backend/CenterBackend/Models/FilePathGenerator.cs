using CenterBackend.IServices;
using CenterBackend.Logging;

namespace CenterBackend.Models
{

    public class FilePathGenerator
    {
        private const string RootDirName = "Report";
        private const string TempDirName = "Temp";

        private readonly IWebHostEnvironment _webHostEnv;
        private readonly string _rootDirectory;

        public  string TempDirectory;
        public FilePathGenerator(IWebHostEnvironment webHostEnv)
        {
            _webHostEnv = webHostEnv ?? throw new ArgumentNullException(nameof(webHostEnv), "Web主机环境不能为空");
            _rootDirectory = Path.Combine(_webHostEnv.ContentRootPath, RootDirName);
            TempDirectory = Path.Combine(_webHostEnv.ContentRootPath, TempDirName);
        }

        public PathAndName GetDay(DateTime fileDate)
        {
            var date = fileDate.Date;
            var pathAndName = new PathAndName
            {
                Directory = Path.Combine(_rootDirectory, "日报表", $"{date.Year}-{date.Month:00}"),
                FileName = $"日报表-{date.Year}-{date.Month:00}-{date.Day:00}.xlsx",
            };
            return pathAndName; 
        }

        public PathAndName GetMonth(DateTime fileDate)
        {
            var date = fileDate.Date;
            var pathAndName = new PathAndName
            {
                Directory = Path.Combine(_rootDirectory, "月报表", $"{date.Year}"),
                FileName = $"月报表-{date.Year}-{date.Month:00}.xlsx",
            };
            return pathAndName;
        }

        public PathAndName GetYear(DateTime fileDate)
        {
            var date = fileDate.Date;
            var pathAndName = new PathAndName
            {
                Directory = Path.Combine(_rootDirectory, "年报表"),
                FileName = $"年报表-{date.Year}.xlsx",
            };
            return pathAndName;
        }

        public PathAndName GetWeek(DateTime fileDate)
        {
            var date = fileDate.Date;
            DateTime weekFirstDay = GetWeekFirstDay(date);  //获取该日期的本周一
            int weekBelongYear = weekFirstDay.Year;             //本周一 归属的年份
            int weekBelongMonth = weekFirstDay.Month;           //本周一 归属的月份
            int weekNumberInYear = GetWeekNumberInYear(weekFirstDay);//该周在当年的周序号

            var pathAndName = new PathAndName
            {
                Directory = Path.Combine(_rootDirectory, "周报表", $"{weekBelongYear}-{weekBelongMonth:00}"),
                FileName = $"周报表-{weekBelongYear}年{weekNumberInYear:00}周.xlsx",
            };
            return pathAndName;
        }

        private static DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = (int)dt.DayOfWeek - (int)DayOfWeek.Monday; //计算与周一的差值
            if (diff < 0) diff += 7;                            //周日处理为-1，补7天
            return dt.AddDays(-diff).Date;
        }

        private static int GetWeekNumberInYear(DateTime dt)
        {
            DateTime weekFirstDay = GetWeekFirstDay(dt);
            DateTime yearFirstDay = new(weekFirstDay.Year, 1, 1);
            DateTime yearFirstMonday = GetWeekFirstDay(yearFirstDay);

            if (yearFirstMonday > new DateTime(weekFirstDay.Year, 1, 4))// 处理跨年周：如果当年第一个周一在1月4日之后，第一周算下一年
            {
                weekFirstDay = GetWeekFirstDay(new DateTime(weekFirstDay.Year - 1, 12, 31));
                return GetWeekNumberInYear(weekFirstDay);
            }

            int daysDiff = (int)(weekFirstDay - yearFirstMonday).TotalDays;// 计算两个周一的间隔天数，除以7得到周数，+1确保从1开始
            int weekNumber = (daysDiff / 7) + 1;

            return Math.Clamp(weekNumber, 1, 53);// 边界校验：避免周数为0或超过53
        }
    }

    public class PathAndName
    {
        public string Directory { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FullPath
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Directory))
                    throw new InvalidOperationException("目录路径不能为空");
                if (string.IsNullOrWhiteSpace(FileName))
                    throw new InvalidOperationException("文件名不能为空");
                return Path.Combine(Directory, FileName);
            }
        }
    }
}
