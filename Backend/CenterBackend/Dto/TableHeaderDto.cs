namespace CenterBackend.Dto
{
    /// <summary>
    /// 表格表头DTO
    /// </summary>
    public class TableHeaderDto
    {
        public string Prop { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;
    }

    /// <summary>
    /// 小时数据DTO
    /// </summary>
    public class HourDataDto
    {
        public int Hour { get; set; }
        public string Date { get; set; } = string.Empty;
        public bool IsNextDay { get; set; }
        // 动态字段（cell1-cell29），用字典适配灵活字段
        public Dictionary<string, string> Cells { get; set; } = new Dictionary<string, string>();

        // 适配前端直接取值（如row.cell1），需转换为动态对象/扩展属性
        // 也可直接定义cell1-cell29属性，示例用字典兼顾灵活性
        public string this[string key]
        {
            get => Cells.TryGetValue(key, out var val) ? val : string.Empty;
            set => Cells[key] = value;
        }
    }

    /// <summary>
    /// 保存单元格数据请求DTO
    /// </summary>
    public class SaveCellRequestDto
    {
        public int Hour { get; set; }
        public string Date { get; set; } = string.Empty;
        public bool IsNextDay { get; set; }
        public string Prop { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }
}
