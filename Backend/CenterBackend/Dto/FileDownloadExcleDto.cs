namespace CenterBackend.Dto
{
    public class FileDownloadExcleDto
    {
        public int Type { get; set; }
        public DateTime Time { get; set; }
    }
    public class FileDownloadZIPDto
    {
        public int Type { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
    }
}
