namespace CenterBackend.Dto
{
    public class FileDownloadExcelDto
    {
        public int type { get; set; }
        public DateTime Time { get; set; }
    }
    public class FileDownloadZIPDto
    {
        public int type { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
    }
}
