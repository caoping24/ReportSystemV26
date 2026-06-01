using CenterReport.Repository.Models;
using ReportServer.Models;
using ReportServer.Services.IUserService;
using System.Reflection;

namespace ReportServer.Services.UserService
{
    public class TagDataConverter : ITagDataConverter
    {
        public SourceData? ConvertTagsToSourceData(List<TagMap>? tags)
        {

            if (tags == null || tags.Count == 0)
            {
                return null;
            }
            SourceData sourceData = new()
            {
                ReportedTime = DateTime.Now
            };
            try
            {
                // 反射缓存：一次性获取SourceData的所有属性
                Type sourceDataType = typeof(SourceData);
                Dictionary<string, PropertyInfo> floatFieldMap;
                floatFieldMap = sourceDataType.GetProperties()
                                                .Where(prop => prop.PropertyType == typeof(float?)) // 仅筛选float?类型字段
                                                .ToDictionary(
                                                    prop => prop.Name.ToLower(), // 转小写匹配cell3/cell4
                                                    prop => prop,
                                                    StringComparer.OrdinalIgnoreCase
                                                );
                foreach (var tag in tags)
                {
                    try
                    {
                        string dbFieldName = tag.DbFieldName?.Trim() ?? string.Empty;//去空格 判null
                        if (string.IsNullOrEmpty(dbFieldName))
                        {
                            //System.Diagnostics.Debug.WriteLine("跳过空的DbFieldName标签");
                            continue;
                        }
                        if (!floatFieldMap.TryGetValue(dbFieldName, out var targetProperty))//查找当前对应的cell
                        {
                            continue;
                        }

                        if (tag.TagValue is float)
                        {
                            var floatValue = (float?)tag.TagValue;
                            targetProperty.SetValue(sourceData, floatValue);
                        }
                    }

                    catch
                    {
                        continue;
                    }
                }
            }
            catch
            {
            }
            return sourceData;
        }
    }
}
