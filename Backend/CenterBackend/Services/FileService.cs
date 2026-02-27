using CenterBackend.IServices;
using CenterBackend.Logging;
using ICSharpCode.SharpZipLib.Zip;


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

        // 下载单个文件
        public (string FilePath, string ContentType, string DownloadFileName) DownloadFileInfo(string sourceFileDirectory, string downloadFileName)
        {
            // 检查参数是否为空
            if (string.IsNullOrWhiteSpace(sourceFileDirectory))
            {
                throw new ArgumentNullException(nameof(sourceFileDirectory));
            }
            if (string.IsNullOrWhiteSpace(downloadFileName))
            {
                throw new ArgumentNullException(nameof(downloadFileName));
            }

            var filePath = Path.Combine(sourceFileDirectory, downloadFileName);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException(nameof(filePath), "要下载的文件不存在");
            }

            return (
                 FilePath: filePath,
                 ContentType: GetMimeType(filePath),
                 DownloadFileName: downloadFileName
             );
        }

        // 将文件夹压缩为Zip包
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
                return File.Exists(zipSavePath);
            }
            catch
            {
                return false;
            }
        }

        // 将单个文件压缩为Zip包
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
                return File.Exists(zipSavePath);
            }
            catch
            {
                return false;
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
                if (!File.Exists(sourceFilePath))
                    return false;

                string? targetFolder = Path.GetDirectoryName(targetFilePath);//提取目标文件夹路径
                if (string.IsNullOrWhiteSpace(targetFolder))
                    return false;

                CreateFolder(targetFolder);//创建目标文件夹
                File.Copy(sourceFilePath, targetFilePath, overwrite);
                return true;
            }
            catch
            {
                return false;
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

        private static string GetMimeType(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".txt" => "text/plain",
                ".pdf" => "application/pdf",
                ".doc" => "application/msword",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xls" => "application/vnd.ms-excel",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".zip" => "application/zip",
                ".rar" => "application/x-rar-compressed",
                ".csv" => "text/csv",
                _ => "application/octet-stream", // 默认二进制流
            };
        }

    }
}