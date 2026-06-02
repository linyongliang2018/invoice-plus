package org.example.invoice;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Stream;

/**
 * 批量将目录中的 PDF 转为图片：按“人名”分组，并把同一份 PDF 的多页拼成一张长图。
 */
public class PdfToImageBatchApp {

    /** PDF 输入目录（会递归扫描 .pdf），按需修改 */
    private static final Path PDF_DIR = Path.of("C:\\Users\\40696\\Desktop\\pdfs");

    /** 图片输出根目录，按需修改 */
    private static final Path OUTPUT_DIR = Path.of("C:\\Users\\40696\\Desktop\\pdf-images");

    /** 渲染分辨率（DPI） */
    private static final float DPI = 200f;

    /** 输出图片格式（建议 png） */
    private static final String IMAGE_FORMAT = "png";

    /** 文本中未命中指定人员关键字时的目录名 */
    private static final String UNMATCHED_PERSON_DIR = "未匹配人员";

    /**
     * 人员分组规则（按顺序匹配）：
     * key = 输出目录名，value = 该人可能出现的名称/别名关键词。
     * 新增人时，按下面格式继续 put 即可。
     */
    private static final Map<String, List<String>> PERSON_GROUPS = new LinkedHashMap<>();

    static {
        PERSON_GROUPS.put("谭惠仪", List.of("谭惠仪"));
        PERSON_GROUPS.put("林儒勋", List.of("林儒勋", "林儒助", "林倩勋"));
        // 示例：PERSON_GROUPS.put("张三", List.of("张三", "张叁"));
    }

    public static void main(String[] args) throws Exception {
        List<Path> pdfFiles = collectPdfPaths(PDF_DIR);
        if (pdfFiles.isEmpty()) {
            System.err.println("未找到任何 PDF 文件：" + PDF_DIR);
            System.exit(1);
        }

        Files.createDirectories(OUTPUT_DIR);

        int convertedPdf = 0;
        int totalPages = 0;
        for (Path pdf : pdfFiles) {
            ConvertResult result = convertSinglePdf(pdf, OUTPUT_DIR);
            convertedPdf++;
            totalPages += result.pages();
            System.out.printf("已转换：%s（分组：%s，%d 页）%n", pdf.getFileName(), result.groupName(), result.pages());
        }

        System.out.printf("完成：共转换 %d 个 PDF，输出 %d 张 %s 图片。输出目录：%s%n",
                convertedPdf, totalPages, IMAGE_FORMAT.toUpperCase(Locale.ROOT), OUTPUT_DIR.toAbsolutePath());
    }

    private static List<Path> collectPdfPaths(Path path) throws Exception {
        List<Path> pdfFiles = new ArrayList<>();
        if (!Files.exists(path)) {
            System.err.println("路径不存在: " + path);
            return pdfFiles;
        }

        if (Files.isDirectory(path)) {
            try (Stream<Path> stream = Files.walk(path)) {
                stream.filter(p -> Files.isRegularFile(p)
                                && p.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".pdf"))
                        .sorted()
                        .forEach(pdfFiles::add);
            }
        } else if (path.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".pdf")) {
            pdfFiles.add(path);
        } else {
            System.err.println("不是 PDF 文件: " + path);
        }
        return pdfFiles;
    }

    private static ConvertResult convertSinglePdf(Path pdfPath, Path outputRoot) throws Exception {
        String baseName = stripExtension(pdfPath.getFileName().toString());
        try (PDDocument document = Loader.loadPDF(pdfPath.toFile())) {
            String pdfText = extractPdfText(document);
            String groupName = resolveGroupByKeywords(pdfText);
            String safeGroupDir = sanitizeAsDirName(groupName);
            Path outDir = outputRoot.resolve(safeGroupDir); // 只按人名分组，一层目录即可
            Files.createDirectories(outDir);

            PDFRenderer renderer = new PDFRenderer(document);
            int pages = document.getNumberOfPages();
            List<BufferedImage> renderedPages = new ArrayList<>(pages);
            for (int i = 0; i < pages; i++) {
                renderedPages.add(renderer.renderImageWithDPI(i, DPI, ImageType.RGB));
            }

            BufferedImage longImage = stitchVertically(renderedPages);
            Path outFile = outDir.resolve(baseName + "." + IMAGE_FORMAT);
            ImageIO.write(longImage, IMAGE_FORMAT, outFile.toFile());

            return new ConvertResult(safeGroupDir, pages);
        }
    }

    private static BufferedImage stitchVertically(List<BufferedImage> images) {
        if (images == null || images.isEmpty()) {
            throw new IllegalArgumentException("images is empty");
        }

        // 统一宽度（用第一张图宽度），避免不同页宽导致错位
        int targetWidth = images.get(0).getWidth();

        long totalHeightLong = 0L;
        int[] scaledHeights = new int[images.size()];
        for (int i = 0; i < images.size(); i++) {
            BufferedImage img = images.get(i);
            int scaledHeight = (int) Math.round((double) img.getHeight() * targetWidth / img.getWidth());
            scaledHeights[i] = scaledHeight;
            totalHeightLong += scaledHeight;
        }
        if (totalHeightLong > Integer.MAX_VALUE) {
            throw new IllegalStateException("拼接后的图片高度过大，无法创建 BufferedImage: " + totalHeightLong);
        }

        int totalHeight = (int) totalHeightLong;
        BufferedImage stitched = new BufferedImage(targetWidth, totalHeight, BufferedImage.TYPE_INT_RGB);

        Graphics2D g = stitched.createGraphics();
        try {
            g.setColor(Color.WHITE);
            g.fillRect(0, 0, targetWidth, totalHeight);
            g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
            g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);

            int y = 0;
            for (int i = 0; i < images.size(); i++) {
                BufferedImage img = images.get(i);
                int scaledHeight = scaledHeights[i];
                // 按统一宽度缩放，然后纵向拼接
                g.drawImage(img, 0, y, targetWidth, scaledHeight, null);
                y += scaledHeight;
            }
        } finally {
            g.dispose();
        }

        return stitched;
    }

    private static String extractPdfText(PDDocument document) throws Exception {
        PDFTextStripper stripper = new PDFTextStripper();
        stripper.setStartPage(1);
        stripper.setEndPage(document.getNumberOfPages());
        String text = stripper.getText(document);
        return text == null ? "" : text.replace('\u3000', ' ');
    }

    private static String resolveGroupByKeywords(String pdfText) {
        if (pdfText == null || pdfText.isBlank()) {
            return UNMATCHED_PERSON_DIR;
        }

        String normalized = pdfText.replaceAll("\\s+", "");
        for (Map.Entry<String, List<String>> entry : PERSON_GROUPS.entrySet()) {
            for (String keyword : entry.getValue()) {
                String normalizedKeyword = keyword.replaceAll("\\s+", "");
                if (!normalizedKeyword.isBlank() && normalized.contains(normalizedKeyword)) {
                    return entry.getKey();
                }
            }
        }
        return UNMATCHED_PERSON_DIR;
    }

    private static String sanitizeAsDirName(String name) {
        String safe = name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
        return safe.isBlank() ? UNMATCHED_PERSON_DIR : safe;
    }

    private static String stripExtension(String filename) {
        int dot = filename.lastIndexOf('.');
        return dot <= 0 ? filename : filename.substring(0, dot);
    }

    private record ConvertResult(String groupName, int pages) {
    }
}
