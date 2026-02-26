namespace CenterBackend.Dto
{
    //传入日期类型，计算数据并记录  
    public class CalculateAndInsertDto
    {
        public int Type { get; set; }
        public DateTime Time { get; set; }
    }
}
