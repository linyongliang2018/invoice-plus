package org.example.invoice;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.stream.Stream;

/** 批量将 PDF 转为图片，每页输出一张 PNG。 */
public class SimplePdfToImageBatchApp {

    /** PDF 输入目录，会递归扫描子目录，按需修改。 */
    private static final Path PDF_DIR = Path.of("C:\\Users\\40696\\Desktop\\医院报销");

    /** 图片输出目录，按需修改。 */
    private static final Path OUTPUT_DIR = PDF_DIR.resolve("图片");

    private static final float DPI = 200f;

    public static void main(String[] args) throws IOException {
        if (!Files.isDirectory(PDF_DIR)) {
            throw new IOException("PDF 输入目录不存在或不是目录：" + PDF_DIR);
        }

        List<Path> pdfFiles;
        try (Stream<Path> paths = Files.walk(PDF_DIR)) {
            pdfFiles = paths.filter(Files::isRegularFile)
                    .filter(path -> path.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".pdf"))
                    .sorted()
                    .toList();
        }
        if (pdfFiles.isEmpty()) {
            System.out.println("未找到 PDF 文件：" + PDF_DIR);
            return;
        }

        int converted = 0;
        int totalImages = 0;
        for (Path pdf : pdfFiles) {
            try {
                int pages = convertPdf(pdf, OUTPUT_DIR);
                converted++;
                totalImages += pages;
                System.out.printf("已转换：%s（%d 张图片）%n", pdf, pages);
            } catch (IOException e) {
                System.err.printf("转换失败：%s，原因：%s%n", pdf, e.getMessage());
            }
        }
        System.out.printf("完成：成功 %d 个 PDF，失败 %d 个，成功 PDF 共输出 %d 张 PNG。输出目录：%s%n",
                converted, pdfFiles.size() - converted, totalImages, OUTPUT_DIR.toAbsolutePath());
    }

    private static int convertPdf(Path pdf, Path outputDir) throws IOException {
        String filename = pdf.getFileName().toString();
        String baseName = filename.substring(0, filename.lastIndexOf('.'));
        try (PDDocument document = Loader.loadPDF(pdf.toFile())) {
            Files.createDirectories(outputDir);
            PDFRenderer renderer = new PDFRenderer(document);
            int pages = document.getNumberOfPages();
            for (int page = 0; page < pages; page++) {
                BufferedImage image = renderer.renderImageWithDPI(page, DPI, ImageType.RGB);
                try {
                    Path outputFile = outputDir.resolve(String.format(Locale.ROOT, "%s-page-%03d.png", baseName, page + 1));
                    if (!ImageIO.write(image, "png", outputFile.toFile())) {
                        throw new IOException("无法写入 PNG 图片：" + outputFile);
                    }
                } finally {
                    image.flush();
                }
            }
            return pages;
        }
    }
}
