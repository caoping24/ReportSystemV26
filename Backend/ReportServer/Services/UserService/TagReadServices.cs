using ReportServer.Models;
using ReportServer.Services.IUserService;
using System.Runtime.InteropServices;
using static ReportServer.Services.UserService.LogServices;//日志服务

namespace ReportServer.Services.UserService
{
    internal class TagReadServices() : ITagReadServices
    {

        public async Task<List<TagMap>> ReadAllTagsAsync()
        {
            CCHMIRUNTIME.HMIRuntime? hmi = null;
            CCHMIRUNTIME.IHMITagSet? tagSet = null;
            try
            {
                //初始化所有变量
                var tagMapManagerInstance = TagMapManager.GetAllTagMaps();
                if (tagMapManagerInstance == null || tagMapManagerInstance.Count == 0)
                    return [];

                hmi = new CCHMIRUNTIME.HMIRuntime();

                //检查连接状态
                var isConnectOk = await GetConnectStatus();
                if (!isConnectOk)
                    return [];

                tagSet = hmi.Tags.CreateTagSet();

                if (tagSet == null)
                    return [];

                foreach (var item in tagMapManagerInstance)
                {
                    try
                    {
                        tagSet.Add(item.TagName);
                    }
                    catch
                    {
                        await AsyncLogHelper.LogWarningAsync($"{item.TagName}:标签添加异常(检查是否有重复).");
                    }
                }

                // 批量读取
                tagSet.Read();

                if (tagSet.LastError != 0)
                    await AsyncLogHelper.LogWarningAsync($"批量读取完成，但存在错误码: {tagSet.LastError}");
                //赋值
                foreach (var item in tagMapManagerInstance)
                {
                    CCHMIRUNTIME.IHMITag? singleTag = null;
                    try
                    {
                        singleTag = tagSet.Item(item.TagName);
                        if (singleTag != null)
                        {
                            if (singleTag.LastError == 0)
                            {
                                var value = singleTag.Value;
                                if (value != null)
                                    item.TagValue = value;
                            }
                        }
                    }
                    catch
                    {
                        await AsyncLogHelper.LogWarningAsync($"{item.TagName}:标签读取异常.");
                        return [];
                    }
                    finally
                    {
                        if (singleTag != null) Marshal.ReleaseComObject(singleTag);
                    }
                }
                //返回修改后的列表
                return tagMapManagerInstance;
            }
            catch
            {
                await AsyncLogHelper.LogErrorAsync($"批量读取标签失败.");
                return [];
            }
            finally
            {
                if (tagSet != null)
                    Marshal.ReleaseComObject(tagSet);
                if (hmi != null)
                    Marshal.ReleaseComObject(hmi);

                // 强制清理 COM  RCW
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        public Task<TagMap?> ReadSingleTagsAsync(string tagName)
        {
            return Task.FromResult<TagMap?>(null);
        }

        public async Task<bool> GetConnectStatus()
        {
            CCHMIRUNTIME.HMIRuntime? hmi = null;
            CCHMIRUNTIME.IHMITagSet? tagSet = null;
            CCHMIRUNTIME.IHMITag? singleTag = null;

            string remotePrefix = RemoteWinccTags.ServerPrefix ?? string.Empty;
            string tagname = RemoteWinccTags.WinccS7ConnectionTagName ?? string.Empty;
            if (string.IsNullOrEmpty(remotePrefix) && string.IsNullOrEmpty(tagname))
            {
                await AsyncLogHelper.LogErrorAsync("TagPerfix或诊断变量为空.");
                return false;
            }
            string remoteTag = remotePrefix + tagname;//需要手动在wincc内部变量中创建并且写入json配置
            try
            {
                hmi = new CCHMIRUNTIME.HMIRuntime();
                if (hmi == null)
                {
                    await AsyncLogHelper.LogErrorAsync("AS连接检查:hmi创建失败.");
                    return false;
                }
                tagSet = hmi.Tags.CreateTagSet();
                if (tagSet == null)
                {
                    await AsyncLogHelper.LogErrorAsync("AS连接检查:tagSet创建失败.");
                    return false;
                }
                tagSet.Add(remoteTag);
                tagSet.Read();
                if (tagSet.LastError != 0)
                {
                    await AsyncLogHelper.LogWarningAsync($"AS连接检查: 批量读取异常，错误码: {remoteTag}");
                    return false;
                }
                singleTag = tagSet.Item(remoteTag);
                if (singleTag == null)
                {
                    await AsyncLogHelper.LogWarningAsync($"AS连接检查:tag读取失败或者为创建:{remoteTag}");
                    return false;
                }
                if (singleTag.LastError != 0)
                {
                    await AsyncLogHelper.LogWarningAsync($"AS连接检查:tag读取失败或者为创建:{remoteTag}");
                    return false;
                }
                var value = singleTag.Value;
                if (value != null && Convert.ToInt32(value) == 1)
                {
                    await AsyncLogHelper.LogInfoAsync("AS连接检查: 连接成功");
                    return true;
                }
                else
                {
                    await AsyncLogHelper.LogWarningAsync($"AS连接检查: 连接失败，状态值: {value}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                await AsyncLogHelper.LogWarningAsync($"{remoteTag}:AS连接检查:执行失败.:{ex}");
                return false;
            }
            finally
            {
                // 统一释放 COM 对象
                if (singleTag != null) Marshal.ReleaseComObject(singleTag);
                if (tagSet != null) Marshal.ReleaseComObject(tagSet);
                if (hmi != null) Marshal.ReleaseComObject(hmi);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

    }
}




