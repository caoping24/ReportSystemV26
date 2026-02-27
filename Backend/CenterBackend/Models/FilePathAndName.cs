namespace CenterBackend.Models
{
    public class FilePathAndName
    {
        public string? DailyFilesPath { get; set; }
        public string? DailyFileName { get; set; }
        public string? DailyFilesFullPath { get; set; }

        public string? MonthlyFilesPath { get; set; }
        public string? MonthlyFileName { get; set; }
        public string? MonthlyFilesFullPath { get; set; }

        public string? WeeklyFilesPath { get; set; }
        public string? WeeklyFileName { get; set; }
        public string? WeeklyFilesFullPath { get; set; }

        public string? YearlyFilesPath { get; set; }
        public string? YearlyFileName { get; set; }
        public string? YearlyFilesFullPath { get; set; }

    }
}
