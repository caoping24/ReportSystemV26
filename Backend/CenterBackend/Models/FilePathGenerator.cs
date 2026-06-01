namespace CenterBackend.Models
{
    public class FilePathGenerator
    {
        private const string RootDirName = "Report";//报表保存目录名称
        private const string TempDirName = "Temp";//ZIP临时文件保存目录名称
        private const string ModelDirName = "Model";//模板文件保存目录名称

        private const string modelFileNameDay = "Model-20260116.xlsx";//各个模板文件的名称
        private const string modelFileNameMonth = "Model-20260317-Month.xlsx";
        private const string modelFileNameYear = "Model-20260317-Year.xlsx";
        private const string modelFileNameWeek = "Model-20251208-Week.xlsx";

        private readonly IWebHostEnvironment _webHostEnv;
        private readonly string _rootDirectory;
        private readonly string _modelFileDirectory;

        public readonly string TempDirectory;

        public FilePathGenerator(IWebHostEnvironment webHostEnv)
        {
            _webHostEnv = webHostEnv ?? throw new ArgumentNullException(nameof(webHostEnv), "Web主机环境不能为空");
            _rootDirectory = Path.Combine(_webHostEnv.ContentRootPath, RootDirName);
            TempDirectory = Path.Combine(_webHostEnv.ContentRootPath, TempDirName);
            _modelFileDirectory = Path.Combine(_webHostEnv.WebRootPath, ModelDirName);
        }

        public PathAndName GetByType(DateTime fileDate, int type)
        {
            PathAndName fileInfo = type switch
            {
                1 => GetDay(fileDate),
                2 => GetMonth(fileDate),
                3 => GetYear(fileDate),
                4 => GetWeek(fileDate),
                _ => new PathAndName(),
            };
            return fileInfo;
        }

        public PathAndName GetDay(DateTime fileDate)
        {
            var date = fileDate.Date;
            var pathAndName = new PathAndName
            {
                Type = 1,
                Directory = Path.Combine(_rootDirectory, "日报表", $"{date.Year}-{date.Month:00}"),
                FileName = $"日报表-{date.Year}-{date.Month:00}-{date.Day:00}.xlsx",

                ModFilePath = Path.Combine(_modelFileDirectory, modelFileNameDay),
            };
            return pathAndName;
        }

        public PathAndName GetMonth(DateTime fileDate)
        {
            var date = fileDate.Date;
            var pathAndName = new PathAndName
            {
                Type = 2,
                Directory = Path.Combine(_rootDirectory, "月报表", $"{date.Year}"),
                FileName = $"月报表-{date.Year}-{date.Month:00}.xlsx",

                ModFilePath = Path.Combine(_modelFileDirectory, modelFileNameMonth),
            };
            return pathAndName;
        }

        public PathAndName GetYear(DateTime fileDate)
        {
            var date = fileDate.Date;
            var pathAndName = new PathAndName
            {
                Type = 3,
                Directory = Path.Combine(_rootDirectory, "年报表"),
                FileName = $"年报表-{date.Year}.xlsx",

                ModFilePath = Path.Combine(_modelFileDirectory, modelFileNameYear),
            };
            return pathAndName;
        }

        public PathAndName GetWeek(DateTime fileDate)
        {
            var date = fileDate.Date;
            DateTime weekFirstDay = GetWeekFirstDay(date);  //每周4 是第一天
            int weekBelongYear = weekFirstDay.Year;             //本周一 归属的年份
            int weekBelongMonth = weekFirstDay.Month;           //本周一 归属的月份
            int weekNumberInYear = GetWeekNumberInYear(weekFirstDay);//该周在当年的周序号

            var pathAndName = new PathAndName
            {
                Type = 4,
                Directory = Path.Combine(_rootDirectory, "周报表", $"{weekBelongYear}-{weekBelongMonth:00}"),
                FileName = $"周报表-{weekBelongYear}年{weekNumberInYear:00}周.xlsx",

                ModFilePath = Path.Combine(_modelFileDirectory, modelFileNameWeek),
            };
            return pathAndName;
        }
        /// <summary>
        /// 每周四是星期的第一天，计算给定日期所在周的第一天（周四）。如果给定日期是周四，则返回该日期；如果是周五，则返回前一天；如果是周三，则返回后一天；以此类推。这样可以确保每周的第一天始终是周四，方便按照周四来归类和命名周报表。
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private static DateTime GetWeekFirstDay(DateTime dt)
        {
            int diff = (int)dt.DayOfWeek - (int)DayOfWeek.Thursday; // 每周的第一天是星期四
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
        public int Type { get; set; } = -1;//默认值表示未设置类型，调用方应检查此值以验证请求参数的有效性
        public DateTime ReportedTime { get; set; }
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
        public string ModFilePath { get; set; } = string.Empty;


    }
}
