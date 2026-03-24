using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ReportServer.Models
{

    public class TagMap(string tagName, string dbFieldName) : INotifyPropertyChanged
    {
        private string _tagName = tagName;
        private string _dbFieldName = dbFieldName;
        private object? _tagValue;
        public string TagName// 标签名
        {
            get => _tagName;
            set
            {
                _tagName = value;
                OnPropertyChanged();
            }
        }
        public string DbFieldName// 数据库字段名
        {
            get => _dbFieldName;
            set
            {
                _dbFieldName = value;
                OnPropertyChanged();
            }
        }
        public object? TagValue// 标签值
        {
            get => _tagValue;
            set
            {
                _tagValue = value;
                OnPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null!)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// 全局映射关系管理器（静态类，统一存储所有映射）
    /// </summary>
    internal static class TagMapManager
    {
        private static readonly List<TagMap> _allTagMaps = []; // 静态只读集合，初始化后全局共享
        /// <summary>
        /// 静态构造函数：在这里集中初始化所有映射关系
        /// </summary>
        /// 
        static TagMapManager()
        {
            AddTagMap("LT04441/MonAnalog.PV_Out#Value", "cell1");
            AddTagMap("TT04751/MonAnalog.PV_Out#Value", "cell2");
            AddTagMap("AT04112/MonAnalog.PV_Out#Value", "cell3");
            AddTagMap("FT04111/FQ.Out#Value", "cell4");
            AddTagMap("FT04113/FQ.Out#Value", "cell5");
            AddTagMap("AT04212/MonAnalog.PV_Out#Value", "cell6");
            AddTagMap("FT05112/质量流量.PV_Out#Value", "cell7");
            AddTagMap("FT05112/FQ.Out#Value", "cell8");
            AddTagMap("TT05151/MonAnalog.PV_Out#Value", "cell9");
            AddTagMap("PT05115/MonAnalog.PV_Out#Value", "cell10");
            AddTagMap("TT05121/MonAnalog.PV_Out#Value", "cell11");
            AddTagMap("TT05131/MonAnalog.PV_Out#Value", "cell12");
            AddTagMap("FT05141/质量流量.PV_Out#Value", "cell13");
            AddTagMap("FT05141/FQ.Out#Value", "cell14");
            AddTagMap("FT05142/质量流量.PV_Out#Value", "cell15");
            AddTagMap("FT05142/FQ.Out#Value", "cell16");
            AddTagMap("TT05143/MonAnalog.PV_Out#Value", "cell17");
            AddTagMap("PT05212/MonAnalog.PV_Out#Value", "cell18");
            AddTagMap("FT05213/体积流量.PV_Out#Value", "cell19");
            AddTagMap("FT05213/FQ.Out#Value", "cell20");
            AddTagMap("TT05214/MonAnalog.PV_Out#Value", "cell21");
            AddTagMap("AT05140/MonAnalog.PV_Out#Value", "cell22");
            AddTagMap("S7$程序(1)/RecordSystem.Cell23", "cell23");
            AddTagMap("PT05311/MonAnalog.PV_Out#Value", "cell24");
            AddTagMap("TT05312/MonAnalog.PV_Out#Value", "cell25");
            AddTagMap("TT05300/MonAnalog.PV_Out#Value", "cell26");
            AddTagMap("PT05317/MonAnalog.PV_Out#Value", "cell27");
            AddTagMap("TT05316/MonAnalog.PV_Out#Value", "cell28");
            //cell29人工录入
            //cell30人工录入
            //cell31人工录入
            //cell32人工录入
            //cell33人工录入
            //cell34人工录入
            //cell35人工录入
            AddTagMap("FT05411/MonAnalog.PV_Out#Value", "cell36");
            AddTagMap("FT05411/FQ.Out#Value", "cell37");
            AddTagMap("LT05417/MonAnalog.PV_Out#Value", "cell38");
            AddTagMap("TT05324/MonAnalog.PV_Out#Value", "cell39");
            AddTagMap("PT05327/MonAnalog.PV_Out#Value", "cell40");
            AddTagMap("FT05323/MonAnalog.PV_Out#Value", "cell41");
            AddTagMap("FT05323/FQ.Out#Value", "cell42");
            //cell43预留
            //cell44预留
            //cell45预留
            //cell46预留
            //cell47预留
            //cell48预留
            //cell49预留
            //cell50预留
            AddTagMap("PT05327/MonAnalog.PV_Out#Value", "cell51");
            AddTagMap("PT05332/MonAnalog.PV_Out#Value", "cell52");
            AddTagMap("LT05344/MonAnalog.PV_Out#Value", "cell53");
            AddTagMap("FT05343/MonAnalog.PV_Out#Value", "cell54");
            AddTagMap("FT05343/FQ.Out#Value", "cell55");
            //cell56人工录入
            //cell57人工录入
            //cell58人工录入
            //cell59人工录入
            //cell60人工录入
            AddTagMap("LT06945/MonAnalog.PV_Out#Value", "cell61");
            AddTagMap("TT06971/MonAnalog.PV_Out#Value", "cell62");
            AddTagMap("SIC06P91/PID.MV_ChnST#Value", "cell63");
            AddTagMap("SIC06P92/PID.MV_ChnST#Value", "cell64");
            AddTagMap("AT06111/MonAnalog.PV_Out#Value", "cell65");
            AddTagMap("TT06171/MonAnalog.PV_Out#Value", "cell66");
            AddTagMap("LT06181/MonAnalog.PV_Out#Value", "cell67");
            AddTagMap("SIC06P01/PID.MV_ChnST#Value", "cell68");
            AddTagMap("FT06122/MonAnalog.PV_Out#Value", "cell69");
            AddTagMap("TT06125/MonAnalog.PV_Out#Value", "cell70");
            AddTagMap("TT06126/MonAnalog.PV_Out#Value", "cell71");
            AddTagMap("SIC06P03/PID.MV_ChnST#Value", "cell72");
            AddTagMap("FT06132/MonAnalog.PV_Out#Value", "cell73");
            AddTagMap("TT06135/MonAnalog.PV_Out#Value", "cell74");
            AddTagMap("TT06136/MonAnalog.PV_Out#Value", "cell75");
            //cell76未安装FT06150-密度
            //cell77未安装FT06150-流量
            AddTagMap("FT06149/MonAnalog.PV_Out#Value", "cell78");
            AddTagMap("FT06149/FQ.Out#Value", "cell79");
            AddTagMap("LT06473/MonAnalog.PV_Out#Value", "cell80");
            AddTagMap("LT06524/MonAnalog.PV_Out#Value", "cell81");
            //cell82人工录入
            //cell83人工录入
            //cell84人工录入
            //cell85人工录入
            //cell86人工录入
            //cell87人工录入
            AddTagMap("TT06311/MonAnalog.PV_Out#Value", "cell88");
            AddTagMap("LT06315/MonAnalog.PV_Out#Value", "cell89");
            AddTagMap("TT07143/MonAnalog.PV_Out#Value", "cell90");
            AddTagMap("LT07144/MonAnalog.PV_Out#Value", "cell91");
            AddTagMap("PT07145/MonAnalog.PV_Out#Value", "cell92");
            //cell93预留
            //cell94预留
            //cell95预留
            //cell96预留
            //cell97预留
            //cell98预留
            //cell99预留
            //cell100预留
            AddTagMap("FT06611/MonAnalog.PV_Out#Value", "cell101");
            AddTagMap("FT06611/FQ.Out#Value", "cell102");
            AddTagMap("TT09171/MonAnalog.PV_Out#Value", "cell103");
            //cell104无-成套系统通讯-脱色液进料流量
            //cell105无-成套系统通讯-脱色液进料累计
            AddTagMap("FT07111/MonAnalog.PV_Out#Value", "cell106");
            AddTagMap("FT07111/FQ.Out#Value", "cell107");
            AddTagMap("TT07152/MonAnalog.PV_Out#Value", "cell108");
            AddTagMap("SIC07P01/PID.MV_ChnST#Value", "cell109");
            AddTagMap("TT07126/MonAnalog.PV_Out#Value", "cell110");
            AddTagMap("FT07126/MonAnalog.PV_Out#Value", "cell111");
            AddTagMap("FT07126/FQ.Out#Value", "cell112");
            AddTagMap("TT07151/MonAnalog.PV_Out#Value", "cell113");
            AddTagMap("FT07137/密度.PV_Out#Value", "cell114");
            AddTagMap("FT07122/MonAnalog.PV_Out#Value", "cell115");
            AddTagMap("FT07122/FQ.Out#Value", "cell116");
            AddTagMap("FT07174/MonAnalog.PV_Out#Value", "cell117");
            AddTagMap("FT07174/FQ1.Out#Value", "cell118");
            AddTagMap("FT07174/MonAnalog.PV_Out#Value", "cell119");
            AddTagMap("FT07174/FQ.Out#Value", "cell120");
            AddTagMap("LT08181/MonAnalog.PV_Out#Value", "cell121");
            AddTagMap("TT08171/MonAnalog.PV_Out#Value", "cell122");
            AddTagMap("SIC08P01/PID.MV_ChnST#Value", "cell123");
            AddTagMap("FT08122/MonAnalog.PV_Out#Value", "cell124");
            AddTagMap("TT06125/MonAnalog.PV_Out#Value", "cell125");
            AddTagMap("TT06126/MonAnalog.PV_Out#Value", "cell126");
            AddTagMap("SIC08P03/PID.MV_ChnST#Value", "cell127");
            AddTagMap("FT08132/MonAnalog.PV_Out#Value", "cell128");
            AddTagMap("TT06135/MonAnalog.PV_Out#Value", "cell129");
            AddTagMap("TT06136/MonAnalog.PV_Out#Value", "cell130");
            //cell131人工录入
            AddTagMap("LT08522/MonAnalog.PV_Out#Value", "cell132");
            AddTagMap("FT08523/MonAnalog.PV_Out#Value", "cell133");
            AddTagMap("FT08523/FQ.Out#Value", "cell134");
            AddTagMap("LT08315/MonAnalog.PV_Out#Value", "cell135");
            AddTagMap("TT08311/MonAnalog.PV_Out#Value", "cell136");
            //cell137人工录入
            //cell138人工录入
            //cell139人工录入
            //cell140人工录入
            //cell141人工录入
            //cell142人工录入
            //cell143预留
            //cell144预留
            //cell145预留
            //cell146预留
            //cell147预留
            //cell148预留
            //cell149预留
            //cell150预留

        }
        /// <summary>
        /// 私有化添加方法，避免外部随意修改映射
        /// </summary>
        private static void AddTagMap(string tagName, string dbFieldName)
        {
            string _serverPrefix = RemoteWinccTags.ServerPrefix;//获取服务器前缀
            if (string.IsNullOrEmpty(_serverPrefix)) //必须获取服务器前缀,如果为空则退出
                return;
            if (!_allTagMaps.Any(t => t.TagName.Equals(tagName, StringComparison.OrdinalIgnoreCase)))
            {
                _allTagMaps.Add(new TagMap(_serverPrefix + tagName, dbFieldName));
            }
        }
        /// <summary>
        /// 获取所有映射关系（只读）
        /// </summary>
        public static List<TagMap> GetAllTagMaps()
        {
            return [.. _allTagMaps]; 
        }
        /// <summary>
        /// 根据变量名查询映射
        /// </summary>
        public static TagMap? GetTagMapByTagName(string tagName)
        {
            return _allTagMaps.FirstOrDefault(t => t.TagName.Equals(tagName, StringComparison.OrdinalIgnoreCase));
        }
    }
}