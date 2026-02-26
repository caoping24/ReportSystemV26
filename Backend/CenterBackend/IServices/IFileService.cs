using CenterBackend.Models;

namespace CenterBackend.IFileService
{
    /// <summary>
    /// 文件操作服务接口
    /// </summary>
    public interface IFileServices
    {
        /// <summary>
        /// 创建文件夹兼容单层/多层，路径存在则跳过，无异常抛出
        /// </summary>
        /// <param name="folderPath">待创建的文件夹完整物理路径</param>
        void CreateFolder(string folderPath);

        /// <summary>
        /// 复制文件（自动创建目标文件夹，解决原生复制无文件夹报错问题）
        /// </summary>
        /// <param name="sourceFilePath">源文件完整物理路径</param>
        /// <param name="targetFilePath">目标文件完整物理路径</param>
        /// <param name="overwrite">是否覆盖已存在的同名目标文件，默认值:True</param>
        /// <returns>复制成功返回True，失败/异常返回False</returns>
        bool CopyFile(string sourceFilePath, string targetFilePath, bool overwrite = true);

        /// <summary>
        /// 判断指定文件是否存在
        /// </summary>
        /// <param name="filePath">文件完整物理路径</param>
        /// <returns>文件存在返回True，不存在/路径为空返回False</returns>
        //bool IsFileExists(string filePath);

        /// <summary>
        /// 判断指定文件夹是否存在
        /// </summary>
        /// <param name="folderPath">文件夹完整物理路径</param>
        /// <returns>文件夹存在返回True，不存在/路径为空返回False</returns>
        //bool IsFolderExists(string folderPath);

        /// <summary>
        /// 删除指定文件，文件不存在则跳过无异常抛出
        /// </summary>
        /// <param name="filePath">待删除的文件完整物理路径</param>
        /// <returns>删除成功返回True，失败/异常返回False</returns>
        //bool DeleteFile(string filePath);

        /// <summary>
        /// 压缩指定文件夹下所有内容至Zip包（含所有子文件夹+文件，保留原层级结构，解决中文乱码）
        /// </summary>
        /// <param name="sourceFolderPath">待压缩的源文件夹完整物理路径</param>
        /// <param name="zipSavePath">压缩包保存的完整物理路径(含文件名+后缀，例:D:\doc\备份2026.zip)</param>
        /// <returns>压缩成功返回True，失败/异常返回False</returns>
        bool CompressFolderToZip(string sourceFolderPath, string zipSavePath);

        /// <summary>
        /// 压缩单个文件至独立Zip包（自动创建目标文件夹，解决中文文件名乱码）
        /// </summary>
        /// <param name="sourceFilePath">待压缩的源文件完整物理路径</param>
        /// <param name="zipSavePath">压缩包保存的完整物理路径(含文件名+后缀)</param>
        /// <returns>压缩成功返回True，失败/异常返回False</returns>
        bool CompressSingleFileToZip(string sourceFilePath, string zipSavePath);
        /// <summary>
        ///直接下载单个文件
        /// </summary>
        (FileStream? stream, string? encodeFileName) DownloadSingleFile(string sourceFilePath, string fileName);
        /// <summary>
        /// 获取指定日期的文件路径+文件名（日报/周报/年报）
        /// 日报按目标日期年月 | 周报按本周一的年月+年周序号 | 年报按年份
        /// </summary>rns>指定层级的完整物理路径，非法类型/空路径返回空字符串</returns>
        FilePathAndName GetDateFolderPathAndName(string rootPath, DateTime targetDate);


        /// <summary>
        /// 按日期创建年/月/周三级文件夹(存在则不创建)
        /// </summary>
        /// <param name="rootPath">存储根路径</param>
        /// <param name="targetDate">目标日期</param>
        void CreateDateFolder(string rootPath, DateTime targetDate);
    }
}