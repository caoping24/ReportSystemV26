using CenterBackend.IServices;
using CenterBackend.Logging;
using CenterBackend.Models;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.AspNetCore.Mvc;
using System.Web;
namespace CenterBackend.Services
{
    /// <summary>
    /// 文件操作服务实现类 - 实现所有文件相关的具体逻辑
    /// </summary>
    public class FileService : IFileServices
    {
        private readonly IAppLogger _logger;
        private readonly IWebHostEnvironment _webHostEnv;
        public FileService(IAppLogger IAppLogger, IWebHostEnvironment webHostEnv)
        {
            this._logger = IAppLogger;
            this._webHostEnv = webHostEnv;
        }

        // 文件夹压缩包下载(压缩指定文件夹内所有内容为Zip包含子文件夹+保持原结构)
        public bool CompressFolderToZip(string sourceFolderDirectory, string zipSaveDirectory, string zipFileName)
        {
            try
            {
                if (!Directory.Exists(sourceFolderDirectory)) throw new DirectoryNotFoundException($"文件夹不存在：{sourceFolderDirectory}");//校验源文件夹

                if (string.IsNullOrWhiteSpace(zipSaveDirectory))
            {
                    return false;
            }
                CreateFolder(zipSaveDirectory);//自动创建存储目录
                string zipSavePath = Path.Combine(zipSaveDirectory, zipFileName);
                using (var fs = new FileStream(zipSavePath, FileMode.Create, FileAccess.Write))
                using (var zipStream = new ZipOutputStream(fs) { IsStreamOwner = true })
                {
                    zipStream.SetLevel(6);//0-9，6=速度+压缩率最优平衡
                    CompressDirectory(sourceFolderDirectory, zipStream, sourceFolderDirectory);//递归压缩文件/子文件夹
                    zipStream.Finish();
                }
                return File.Exists(zipSavePath);//校验压缩包是否生成
            }
            catch
            {
                return false;
            }
        }

        // 单个文件压缩包下载(压缩为独立Zip包)
        public bool CompressSingleFileToZip(string sourceFilePath, string zipSaveDirectory, string zipFileName)
        {
            try
            {
                if (!File.Exists(sourceFilePath)) throw new FileNotFoundException($"文件不存在：{sourceFilePath}");//校验源文件
                if (string.IsNullOrWhiteSpace(zipSaveDirectory))
                {
                    return false;
                }
                CreateFolder(zipSaveDirectory);//自动创建存储目录
                string zipSavePath = Path.Combine(zipSaveDirectory, zipFileName);
                using (var fs = new FileStream(zipSavePath, FileMode.Create, FileAccess.Write))// 流式写入，缓冲区读取，适配大文件
                using (var zipStream = new ZipOutputStream(fs) { IsStreamOwner = true })
                {
                    zipStream.SetLevel(6);//最优压缩级别
                    var fileInfo = new FileInfo(sourceFilePath);
                    var zipEntry = new ZipEntry(fileInfo.Name) { DateTime = DateTime.Now };//创建压缩项
                    zipStream.PutNextEntry(zipEntry);

                    byte[] buffer = new byte[4096];//4K缓冲区，平衡性能与内存
                    using (var fileStream = new FileStream(sourceFilePath, FileMode.Open, FileAccess.Read))
                    {
                        int bytesRead;
                        while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            zipStream.Write(buffer, 0, bytesRead);
                        }
                    }
                    zipStream.CloseEntry();
                    zipStream.Finish();
                }
                return File.Exists(zipSavePath);//校验压缩包是否生成
            }
            catch
            {
                return false;
            }
        }
        // 单个文件直接下载,不压缩
        public (FileStream? stream, string? encodeFileName) DownloadSingleFile(string sourceFileDirectory, string fileName)
        {
            try
            {
                string filePath = Path.Combine(sourceFileDirectory, fileName);
                if (!File.Exists(filePath)) return (null, null);
                var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);// 只读方式打开文件流，共享读取（文件被占用也能下载），流式返回，不加载到内存
                string encodeFileName = HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8);     // 解决中文文件名下载乱码问题
                return (fileStream, encodeFileName);
            }
            catch
            {
                return (null, null);
            }
        }

        // 递归压缩文件夹内所有文件/子文件夹，保持原目录相对结构
        private static void CompressDirectory(string folderPath, ZipOutputStream zipStream, string rootFolder)
        {
            foreach (string file in Directory.GetFiles(folderPath))//遍历当前文件夹所有文件
            {
                var fileInfo = new FileInfo(file);
                string entryName = file[(rootFolder.Length + 1)..];//生成压缩包内相对路径
                var zipEntry = new ZipEntry(entryName) { DateTime = DateTime.Now };
                zipStream.PutNextEntry(zipEntry);

                byte[] buffer = new byte[4096];//4K缓冲区
                using (var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    int bytesRead;
                    while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        zipStream.Write(buffer, 0, bytesRead);
                    }
                }
                zipStream.CloseEntry();
            }

            foreach (string dir in Directory.GetDirectories(folderPath))//递归遍历子文件夹
            {
                CompressDirectory(dir, zipStream, rootFolder);
            }
        }

        // 创建文件夹兼容单层/多层，路径存在则跳过，无异常抛出
        public void CreateFolder(string folderDirectory)
        {
            if (!string.IsNullOrWhiteSpace(folderDirectory) && !Directory.Exists(folderDirectory))//判空+判存在
        {
                Directory.CreateDirectory(folderDirectory);
            }
        }

        // 复制文件(自动创建目标文件夹)
        public bool CopyFile(string sourceFilePath, string targetFilePath, bool overwrite = true)
        {
            try
            {
                if (!File.Exists(sourceFilePath)) throw new FileNotFoundException($"源文件不存在：{sourceFilePath}");//校验源文件

                string? targetFolder = Path.GetDirectoryName(targetFilePath);//提取目标文件夹路径
                if (string.IsNullOrWhiteSpace(targetFolder))
                {
                    return false;
                }

                CreateFolder(targetFolder);//创建目标文件夹
                File.Copy(sourceFilePath, targetFilePath, overwrite);
                return true;
            }
            catch
            {
                return false;//异常则返回失败
            }
        }

    }
}