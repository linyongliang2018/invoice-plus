package org.example.invoice;

import org.example.invoice.model.InvoiceData;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

public class InvoiceImportApp {

    /** Excel 汇总输出路径，按需修改 */
    private static final Path EXCEL_PATH = Path.of("D:\\发票汇总.xlsx");

    /** 发票 XML 所在目录（会导入目录下全部 .xml），按需修改 */
    private static final Path XML_DIR = Path.of("C:\\Users\\40696\\Desktop\\京东发票-京东发票明细");

    public static void main(String[] args) throws Exception {
        Path excelPath = EXCEL_PATH;
        List<Path> xmlFiles = collectXmlPaths(XML_DIR);

        if (xmlFiles.isEmpty()) {
            System.err.println("未找到任何 XML 文件：" + XML_DIR);
            System.exit(1);
        }

        List<InvoiceData> invoices = InvoiceXmlParser.parseAll(xmlFiles);
        InvoiceExcelExporter.ExportResult result =
                new InvoiceExcelExporter().exportInvoices(excelPath, invoices);

        System.out.printf("处理完成：共导出 %d 张（跳过重复/无效 %d 张），已全量覆盖。输出：%s%n",
                result.exported(), result.skipped(), excelPath.toAbsolutePath());
    }

    private static List<Path> collectXmlPaths(Path path) throws Exception {
        List<Path> xmlFiles = new ArrayList<>();
        if (!Files.exists(path)) {
            System.err.println("路径不存在: " + path);
            return xmlFiles;
        }
        if (Files.isDirectory(path)) {
            try (Stream<Path> stream = Files.walk(path)) {
                stream.filter(p -> Files.isRegularFile(p)
                                && p.getFileName().toString().toLowerCase().endsWith(".xml"))
                        .sorted()
                        .forEach(xmlFiles::add);
            }
        } else if (path.getFileName().toString().toLowerCase().endsWith(".xml")) {
            xmlFiles.add(path);
        } else {
            System.err.println("不是 XML 文件: " + path);
        }
        return xmlFiles;
    }
}
