using System.Diagnostics;
using System.Reflection;

namespace InvoiceProcessor
{
    public class SumatraPrinter
    {
        // 释放的目标路径 (放在 AppData 目录下)
        private static readonly string ExePath = Path.Combine(FileSystem.Current.AppDataDirectory, "sumatrapdf.exe");

        /// <summary>
        /// 初始化：检查 exe 是否存在，不存在则从资源中释放
        /// </summary>
        public static async Task InitializeAsync()
        {
            if (File.Exists(ExePath)) return; // 已经存在，无需重复释放

            try
            {
                var assembly = Assembly.GetExecutingAssembly();

                // ⚠️ 注意：资源名称格式通常为 "命名空间.文件夹.文件名"
                // 如果文件直接在项目根目录，就是 "项目名.文件名"
                // 请确保这里的 "InvoiceProcessor.sumatrapdf.exe" 与您的项目名一致
                var resourceName = "InvoiceProcessor.sumatrapdf.exe";

                using var stream = assembly.GetManifestResourceStream(resourceName) ?? throw new Exception($"未找到嵌入资源: {resourceName}。请检查生成操作是否设为 Embedded Resource。");
                using var fileStream = File.Create(ExePath);
                await stream.CopyToAsync(fileStream);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"释放 SumatraPDF 失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 调用 SumatraPDF 进行打印
        /// </summary>
        public static void Print(string pdfPath, string printerName = "")
        {
            if (!File.Exists(ExePath))
            {
                throw new FileNotFoundException("SumatraPDF 核心文件丢失，无法打印。");
            }

            // 构建命令行参数
            // -print-to-default: 打印到默认
            // -print-to "Printer Name": 打印到指定打印机
            // -silent: 不显示界面和错误弹窗
            string args;

            if (string.IsNullOrWhiteSpace(printerName))
            {
                args = $"-print-to-default -silent \"{pdfPath}\"";
            }
            else
            {
                args = $"-print-to \"{printerName}\" -silent \"{pdfPath}\"";
            }

            var psi = new ProcessStartInfo
            {
                FileName = ExePath,
                Arguments = args,
                UseShellExecute = false, // 必须为 false 才能重定向输出（虽然这里不需要）
                CreateNoWindow = true,   // 完全隐藏黑窗口
                WindowStyle = ProcessWindowStyle.Hidden
            };

            try
            {
                using var process = Process.Start(psi);
                // SumatraPDF 打印通常是异步指令，启动后很快就会退出
                // 可以选择等待几秒，或者直接通过
                process.WaitForExit(10000); // 最多等待10秒，防止卡死
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"打印命令发送失败: {ex.Message}");
                throw;
            }
        }
    }
}