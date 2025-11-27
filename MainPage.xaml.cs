using System.Diagnostics;

namespace InvoiceProcessor
{
    public partial class MainPage : ContentPage
    {
        private string _sourceFolderPath = string.Empty;
        private string _outputFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ProcessedInvoices");
        private readonly ProcessLogic _logic = new();

        public MainPage()
        {
            InitializeComponent();
            _ = InitializePrinterEngine();
        }

        private async Task InitializePrinterEngine()
        {
            try
            {
                await SumatraPrinter.InitializeAsync();
                Log("打印引擎初始化就绪 (SumatraPDF)");
            }
            catch (Exception ex)
            {
                Log($"[严重错误] 打印引擎初始化失败: {ex.Message}");
                await DisplayAlertAsync("错误", "无法加载内置打印组件，打印功能将不可用。", "OK");
            }
        }

        private void Log(string message)
        {
            // 在主线程更新 UI (MAUI要求)
            MainThread.BeginInvokeOnMainThread(() =>
            {
                LogEditor.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n";
                // 自动滚动到底部
            });
        }

        // 按钮点击事件：选择文件夹
        private async void OnSelectFolderClicked(object sender, EventArgs e)
        {
            try
            {
                string? selectedPath = null;

                // 1. 创建 Windows 专用的文件夹选择器
                var folderPicker = new Windows.Storage.Pickers.FolderPicker();

                // 2. 设置为选取文件夹模式 (必须添加一个通配符过滤器，否则会报错)
                folderPicker.FileTypeFilter.Add("*");

                // 3. 【关键步骤】获取当前 MAUI 窗口的句柄 (HWND)
                // WinUI 3 安全限制要求必须显式关联窗口，否则 Picker 无法弹出
                var window = Application.Current.Windows.FirstOrDefault(w => w.Page == this) ?? App.Current.Windows.FirstOrDefault();
                if (window != null && window.Handler.PlatformView is MauiWinUIWindow winuiWindow)
                {
                    var hwnd = winuiWindow.WindowHandle;
                    // 初始化 Picker 与窗口的关联
                    WinRT.Interop.InitializeWithWindow.Initialize(folderPicker, hwnd);
                }

                // 4. 显示选择器并等待用户操作
                var result = await folderPicker.PickSingleFolderAsync();

                if (result != null)
                {
                    selectedPath = result.Path;
                }

                // 5. 处理结果
                if (!string.IsNullOrEmpty(selectedPath))
                {
                    _sourceFolderPath = selectedPath;
                    SourceFolderEntry.Text = _sourceFolderPath;
                    Log($"已选择源文件夹: {_sourceFolderPath}");

                    // 统计文件数量
                    var pdfFiles = Directory.GetFiles(_sourceFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);
                    FileCountLabel.Text = $"找到 {pdfFiles.Length} 个 PDF 文件。";
                }
                else
                {
                    Log("用户取消了文件夹选择。");
                }
            }
            catch (Exception ex)
            {
                Log($"选择文件夹时出错: {ex.Message}");
                await DisplayAlertAsync("错误", $"选择文件夹失败: {ex.Message}", "确定");
            }
        }
        // 按钮点击事件：开始处理并打印
        private async void OnProcessAndPrintClicked(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_sourceFolderPath) || !Directory.Exists(_sourceFolderPath))
            {
                await DisplayAlertAsync("错误", "请先选择有效的源文件夹。", "确定");
                return;
            }

            LogEditor.Text = string.Empty;
            Log(">>> 开始智能发票处理和打印 <<<");

            // 1. 设置输出文件夹
            if (Directory.Exists(_outputFolderPath))
            {
                Log($"检测到旧的输出文件夹，正在清空: {_outputFolderPath}");
                try
                {
                    // 删除文件夹内容
                    foreach (var file in Directory.GetFiles(_outputFolderPath))
                    {
                        File.Delete(file);
                    }
                }
                catch (Exception ex)
                {
                    Log($"错误：无法清空输出文件夹。请检查文件是否被占用。{ex.Message}");
                    return;
                }
            }
            else
            {
                Directory.CreateDirectory(_outputFolderPath);
            }
            Log($"已准备输出文件夹: {_outputFolderPath}");

            // 2. 获取并排序文件
            var pdfPaths = Directory.GetFiles(_sourceFolderPath, "*.pdf", SearchOption.TopDirectoryOnly).ToList();
            if (!pdfPaths.Any())
            {
                Log("文件夹中没有找到任何 PDF 文件。");
                return;
            }

            // 获取元数据并过滤掉无法读取的文件
            var pdfMetadatas = pdfPaths
                .Select(p => _logic.GetPdfMetadata(p))
                .Where(m => m != null)
                .ToList();

            // 按高度降序排列
            var sortedPdfs = pdfMetadatas.OrderByDescending(m => m.Height).ToList();

            // 3. 执行配对和合并逻辑 (与 Python 脚本的主逻辑相同)
            var filesToPrint = new List<string>();
            int mergedPairCount = 0;
            int standaloneFileCount = 0;

            // 奇数处理
            if (sortedPdfs.Count % 2 != 0)
            {
                Log("\n文件个数为奇数，先单独处理最长 PDF 文件：");
                var longestPdf = sortedPdfs.First();
                sortedPdfs.RemoveAt(0);

                var outputFilename = $"single_{Path.GetFileNameWithoutExtension(longestPdf.Path)}.pdf";
                var outputPath = Path.Combine(_outputFolderPath, outputFilename);
                if (_logic.CreateStandalonePdf(longestPdf.Path, outputPath))
                {
                    filesToPrint.Add(outputPath);
                    standaloneFileCount++;
                }
            }

            // 配对和合并循环
            while (sortedPdfs.Count >= 2)
            {
                var longPdf = sortedPdfs.First();
                var shortPdf = sortedPdfs.Last();
                sortedPdfs.RemoveAt(0);
                sortedPdfs.RemoveAt(sortedPdfs.Count - 1);

                Log($"\n正在尝试配对:\n[L] {Path.GetFileName(longPdf.Path)} \n[S] {Path.GetFileName(shortPdf.Path)}");

                var (l_scale, s_scale) = _logic.FindOptimalScales(longPdf, shortPdf);

                if (l_scale.HasValue && s_scale.HasValue)
                {
                    Log($"  > 找到可合并缩放值: [L] {l_scale.Value * 100:0.0}% | [S] {s_scale.Value * 100:0.0}%");
                    var outputFilename = $"{Path.GetFileNameWithoutExtension(longPdf.Path)}_{Path.GetFileNameWithoutExtension(shortPdf.Path)}.pdf";
                    var outputPath = Path.Combine(_outputFolderPath, outputFilename);
                    if (_logic.CreateMergedPdf(longPdf.Path, l_scale.Value, shortPdf.Path, s_scale.Value, outputPath))
                    {
                        mergedPairCount++;
                        filesToPrint.Add(outputPath);
                    }
                }
                else
                {
                    // 无可配对文件，单独处理长PDF
                    Log($"  > 无法配对。单独处理 {Path.GetFileName(longPdf.Path)}");
                    var outputFilename = $"single_long_{Path.GetFileNameWithoutExtension(longPdf.Path)}.pdf";
                    var outputPath = Path.Combine(_outputFolderPath, outputFilename);
                    if (_logic.CreateStandalonePdf(longPdf.Path, outputPath))
                    {
                        filesToPrint.Add(outputPath);
                        standaloneFileCount++;
                    }

                    // 将短 PDF 放回并重新排序
                    Log($"  > 将较短的 PDF '{Path.GetFileName(shortPdf.Path)}' 放回列表以待下次配对");
                    sortedPdfs.Add(shortPdf);
                    sortedPdfs = sortedPdfs.OrderByDescending(m => m.Height).ToList();
                }
            }

            // 剩余文件处理
            if (sortedPdfs.Any())
            {
                var remainingPdf = sortedPdfs.First();
                Log($"\n--- 处理剩余文件: {Path.GetFileName(remainingPdf.Path)} ---");
                var outputFilename = $"single_rem_{Path.GetFileNameWithoutExtension(remainingPdf.Path)}.pdf";
                var outputPath = Path.Combine(_outputFolderPath, outputFilename);
                if (_logic.CreateStandalonePdf(remainingPdf.Path, outputPath))
                {
                    filesToPrint.Add(outputPath);
                    standaloneFileCount++;
                }
            }

            // 4. 汇总和打印
            Log("\n--- 所有文件处理完毕！---");
            Log($"已配对可合并文件总数: {mergedPairCount * 2} 个");
            Log($"已单独处理文件总数: {standaloneFileCount} 个");
            Log($"总共生成 {filesToPrint.Count} 个目标打印文档。");

            if (filesToPrint.Any())
            {
                var printConfirmed = await DisplayAlertAsync("确认打印", $"准备打印 {filesToPrint.Count} 个目标文档。是否继续发送打印任务？", "是", "否");

                if (printConfirmed)
                {
                    Log("\n--- 开始发送打印任务 ---");
                    foreach (var pdfFile in filesToPrint)
                    {
                        PrintPdf(pdfFile, PrinterNameEntry.Text);
                    }
                    Log("\n所有打印任务已发送完毕！");
                }
                else
                {
                    Log($"打印操作已取消。处理好的文件保存在 '{_outputFolderPath}' 中。");
                }
            }
        }

        // 5. 打印功能 
        private void PrintPdf(string pdfPath, string printerName)
        {
            try
            {
                Log($"正在发送任务: {Path.GetFileName(pdfPath)}");

                // === 使用内置 SumatraPDF 打印 ===
                SumatraPrinter.Print(pdfPath, printerName);

                Log(" -> 指令已发送至打印队列");

                // 稍微延时，防止瞬间大量进程冲击 CPU
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Log($"[错误] 打印失败: {ex.Message}");
            }
        }
    }
}
