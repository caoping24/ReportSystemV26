using CenterBackend.Logging;
using CenterBackend.Models;
using ICSharpCode.SharpZipLib.Zip;
using System.Web;

namespace CenterBackend.IServices
{
    /// <summary>
    /// 文件操作服务接口
    /// </summary>
    public interface IFileServices
    {

        (string FilePath, string ContentType, string DownloadFileName) DownloadFileInfo(string sourceFileDirectory, string downloadFileName);

        // 文件夹压缩包下载(压缩指定文件夹内所有内容为Zip包含子文件夹+保持原结构)
        public bool CompressFolderToZip(string sourceFolderDirectory, string zipSaveDirectory, string zipFileName);

        // 单个文件压缩包下载(压缩为独立Zip包)
        public bool CompressSingleFileToZip(string sourceFilePath, string zipSaveDirectory, string zipFileName);

        // 创建文件夹兼容单层/多层，路径存在则跳过，无异常抛出
        public void CreateFolder(string folderDirectory);
        // 复制文件(自动创建目标文件夹)
        public bool CopyFile(string sourceFilePath, string targetFilePath, bool overwrite = true);
    }
}