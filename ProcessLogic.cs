using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace InvoiceProcessor
{
    internal class ProcessLogic
    {
        // A4尺寸 (单位: point，1 point = 1/72 inch)
        private const double A4_WIDTH = 595;
        private const double A4_HEIGHT = 842;

        // 缩放比例范围 (使用 double 来匹配 C# 的浮点运算)
        private const double SCALE_MIN = 0.70;
        private const double SCALE_MAX = 1.00;
        private const double SCALE_STEP = -0.001;
        private const double DEFAULT_STANDALONE_SCALE = 0.70;

        public class PdfMetadata
        {
            public required string Path { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }
    
        //仅用于获取 PDF 文档的第一页尺寸    
        public PdfMetadata GetPdfMetadata(string pdfPath)
        {
            try
            {
                using var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
                if (document.PageCount == 0) return null;
                var page = document.Pages[0];
                return new PdfMetadata
                {
                    Path = pdfPath,
                    Width = page.Width.Point,
                    Height = page.Height.Point
                };
            }
            catch (Exception ex)
            {
                // 异常处理
                System.Diagnostics.Debug.WriteLine($"无法读取文件 '{Path.GetFileName(pdfPath)}' 的尺寸信息: {ex.Message}");
                return null;
            }
        }

        /// 查找最佳缩放比例
        public (double? LongScale, double? ShortScale) FindOptimalScales(PdfMetadata longPdfMeta, PdfMetadata shortPdfMeta)
        {
            double shortScale = SCALE_MAX;
            while (shortScale >= SCALE_MIN)
            {
                double longScale = SCALE_MAX;
                while (longScale >= SCALE_MIN)
                {
                    // 计算缩放后的高度
                    double scaledLHeight = longPdfMeta.Height * longScale;
                    double scaledSHeight = shortPdfMeta.Height * shortScale;

                    // 检查是否能纵向拼接进A4页面
                    double totalHeight = scaledLHeight + scaledSHeight;

                    if (totalHeight <= A4_HEIGHT)
                    {
                        return (longScale, shortScale);
                    }

                    longScale += SCALE_STEP; // -0.01
                }
                shortScale += SCALE_STEP; // -0.01
            }

            return (null, null); // 未找到合适的缩放比例
        }

        //将两个PDF合并到一页A4上 (使用 PdfSharp)
        public bool CreateMergedPdf(string pdf1Path, double pdf1Scale, string pdf2Path, double pdf2Scale, string outputPath)
        {
            try
            {
                var pdf1Meta = GetPdfMetadata(pdf1Path);
                var pdf2Meta = GetPdfMetadata(pdf2Path);

                using var outputDocument = new PdfDocument();
                // 创建空白A4页面
                var newPage = outputDocument.AddPage();
                newPage.Width = (XUnitPt)A4_WIDTH;
                newPage.Height = (XUnitPt)A4_HEIGHT;

                var gfx = XGraphics.FromPdfPage(newPage);

                // 1. 将第一个PDF (L) 放置在上半部分
                using (var pdf1 = XPdfForm.FromFile(pdf1Path))
                {
                    //默认加载第一页
                    double x1 = (A4_WIDTH - pdf1Meta.Width * pdf1Scale) / 2;
                    double y1 = A4_HEIGHT - pdf1Meta.Height * pdf1Scale; // 顶部对齐
                    double w1 = pdf1Meta.Width * pdf1Scale;
                    double h1 = pdf1Meta.Height * pdf1Scale;

                    gfx.DrawImage(pdf1, new XRect(x1, y1, w1, h1));
                }

                // 2. 将第二个PDF (S) 放置在下半部分
                using (var pdf2 = XPdfForm.FromFile(pdf2Path))
                {
                    double x2 = (A4_WIDTH - pdf2Meta.Width * pdf2Scale) / 2;
                    double y2 = 0; // 底部对齐
                    double w2 = pdf2Meta.Width * pdf2Scale;
                    double h2 = pdf2Meta.Height * pdf2Scale;

                    gfx.DrawImage(pdf2, new XRect(x2, y2, w2, h2));
                }

                outputDocument.Save(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"文件合并失败: {Path.GetFileName(pdf1Path)}, {Path.GetFileName(pdf2Path)}. 错误: {ex.Message}");
                return false;
            }
        }

        //将单个PDF单独放到一页A4上 (使用 PdfSharp)
        public bool CreateStandalonePdf(string singlePdfPath, string outputPath)
        {
            try
            {
                var singlePdfMeta = GetPdfMetadata(singlePdfPath);

                using var outputDocument = new PdfDocument();
                // 创建空白A4页面
                var newPage = outputDocument.AddPage();
                newPage.Width = (XUnitPt)A4_WIDTH;
                newPage.Height = (XUnitPt)A4_HEIGHT;

                var gfx = XGraphics.FromPdfPage(newPage);

                // 放置 PDF
                using (var singlePdf = XPdfForm.FromFile(singlePdfPath))
                {
                    double scale = DEFAULT_STANDALONE_SCALE;
                    double x = (A4_WIDTH - singlePdfMeta.Width * scale) / 2;
                    double y = A4_HEIGHT - singlePdfMeta.Height * scale; // 顶部对齐
                    double w = singlePdfMeta.Width * scale;
                    double h = singlePdfMeta.Height * scale;

                    gfx.DrawImage(singlePdf, new XRect(x, y, w, h));
                }

                outputDocument.Save(outputPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"处理单个文件失败: {Path.GetFileName(singlePdfPath)}. 错误: {ex.Message}");
                return false;
            }
        }
    }
}
