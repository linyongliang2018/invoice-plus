package org.example.invoice;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilderFactory;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 按文件夹全量解析 XML 并导出多 sheet 发票分析报表。
 */
public class InvoiceFolderWorkbookApp {

    private static final Path XML_DIR = Path.of("C:\\Users\\40696\\Desktop\\京东发票-京东发票明细");
    private static final Path OUTPUT_EXCEL = Path.of("D:\\发票整理-多sheet.xlsx");
    private static final DateTimeFormatter ISSUE_TIME_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final LocalDateTime MIN_ISSUE_TIME = LocalDate.parse("2024-06-18").atStartOfDay();

    /**
     * 只影响 sheet4 分组；sheet5 用于展示这些号码在目录中出现的发票。
     */
    private static final List<String> SHEET4_EXCLUDE_INVOICE_NUMBERS = List.of(
             "26337000000291496743",
             "26337000000291496734",
             "26337000000293247755",
             "26337000000291496738",
             "26217000000142792534",
             "26337000000198158888",
             "26467000000026687671",
             "26117000000196262669",
             "26447000000219590875",
             "25127000000508511969",
             "25127000000512743992",
             "25127000000508511971",
             "25227000000101068653",
             "25447000001223139856",
             "25447000001213412083",
             "25447000000727035968",
             "25427000000259180993",
             "25957200000005978372",
            "26457000000095300376",
            "26527000000055014776",
            "26447000000671312987",
            "26337000000318572701",
            "26337000000293083056",
            "26337000000400309359",
            "26957000000123526973",
            "26427000000450539968",
            "26337000000504766983"
    );

    private static final String[] BASE_HEADERS = {
            "开票机构(TaxBureauName)",
            "发票号码(InvoiceNumber)",
            "开票时间(IssueTime)",
            "销方纳税人识别号(SellerIdNum)",
            "销方名称(SellerName)",
            "商品明细(ItemName)",
            "发票价税合计(TotalAmWithoutTax)"
    };

    private static final String[] SHEET4_HEADERS = {
            "分组编号(GroupNo)",
            "分组标签(GroupTag)",
            "开票机构(TaxBureauName)",
            "发票号码(InvoiceNumber)",
            "开票时间(IssueTime)",
            "销方纳税人识别号(SellerIdNum)",
            "销方名称(SellerName)",
            "商品明细(ItemName)",
            "发票价税合计(TotalAmWithoutTax)",
            "分组累计(GroupSum)"
    };

    private static final byte[][] BLOCK_HEADER_COLORS = {
            {(byte) 255, (byte) 230, (byte) 153},
            {(byte) 198, (byte) 224, (byte) 180}
    };

    private static final byte[][] BLOCK_BODY_COLORS = {
            {(byte) 255, (byte) 242, (byte) 204},
            {(byte) 226, (byte) 239, (byte) 218}
    };

    public static void main(String[] args) throws Exception {
        List<Path> xmlFiles = collectXmlFiles(XML_DIR);
        if (xmlFiles.isEmpty()) {
            System.out.println("未找到 XML 文件：" + XML_DIR.toAbsolutePath());
            return;
        }

        List<InvoiceRecord> allInvoices = parseAll(xmlFiles);
        if (allInvoices.isEmpty()) {
            System.out.println("XML 解析完成，但未得到有效发票。");
            return;
        }

        Map<String, InvoiceRecord> unique = new LinkedHashMap<>();
        for (InvoiceRecord invoice : allInvoices) {
            if (!invoice.invoiceNumber().isBlank()) {
                unique.putIfAbsent(invoice.invoiceNumber(), invoice);
            }
        }
        List<InvoiceRecord> invoices = new ArrayList<>(unique.values());
        int beforeDateCount = (int) invoices.stream()
                .filter(it -> safeIssueTime(it).isBefore(MIN_ISSUE_TIME))
                .count();
        invoices = invoices.stream()
                .filter(it -> !safeIssueTime(it).isBefore(MIN_ISSUE_TIME))
                .collect(Collectors.toList());

        List<InvoiceRecord> sheet1 = invoices.stream()
                .filter(it -> it.totalAmWithoutTax().compareTo(new BigDecimal("100")) > 0)
                .sorted(Comparator.comparing(InvoiceFolderWorkbookApp::safeIssueTime).reversed()
                        .thenComparing(InvoiceRecord::invoiceNumber))
                .collect(Collectors.toList());

        Set<String> sheet1ItemSet = sheet1.stream()
                .flatMap(it -> it.items().stream())
                .map(InvoiceFolderWorkbookApp::normItem)
                .filter(s -> !s.isBlank())
                .collect(Collectors.toSet());

        Set<String> sheet1Numbers = sheet1.stream()
                .map(InvoiceRecord::invoiceNumber)
                .collect(Collectors.toSet());

        List<InvoiceRecord> nonSheet1 = invoices.stream()
                .filter(it -> !sheet1Numbers.contains(it.invoiceNumber()))
                .sorted(Comparator.comparing(InvoiceFolderWorkbookApp::safeIssueTime).reversed()
                        .thenComparing(InvoiceRecord::invoiceNumber))
                .collect(Collectors.toList());

        List<InvoiceRecord> sheet2 = new ArrayList<>();
        List<InvoiceRecord> sheet3 = new ArrayList<>();
        for (InvoiceRecord inv : nonSheet1) {
            boolean conflict = inv.items().stream()
                    .map(InvoiceFolderWorkbookApp::normItem)
                    .anyMatch(sheet1ItemSet::contains);
            if (conflict) {
                sheet3.add(inv);
            } else {
                sheet2.add(inv);
            }
        }

        Set<String> excludeSet = new HashSet<>(SHEET4_EXCLUDE_INVOICE_NUMBERS);
        List<InvoiceRecord> sheet5 = invoices.stream()
                .filter(it -> excludeSet.contains(it.invoiceNumber()))
                .sorted(Comparator.comparing(InvoiceFolderWorkbookApp::safeIssueTime).reversed()
                        .thenComparing(InvoiceRecord::invoiceNumber))
                .collect(Collectors.toList());

        List<InvoiceRecord> sheet4Candidates = sheet2.stream()
                .filter(it -> !excludeSet.contains(it.invoiceNumber()))
                .collect(Collectors.toList());

        List<GroupRow> allGroupRows = buildSheet4Groups(sheet4Candidates);
        List<GroupRow> sheet4Rows = allGroupRows.stream()
                .filter(row -> !row.groupTag().startsWith("零散不足100"))
                .collect(Collectors.toList());
        List<InvoiceRecord> sheet6 = allGroupRows.stream()
                .filter(row -> row.groupTag().startsWith("零散不足100"))
                .flatMap(row -> row.invoices().stream())
                .sorted(Comparator.comparing(InvoiceFolderWorkbookApp::safeIssueTime).reversed()
                        .thenComparing(InvoiceRecord::invoiceNumber))
                .collect(Collectors.toList());

        writeWorkbook(sheet1, sheet2, sheet3, sheet4Rows, sheet5, sheet6, OUTPUT_EXCEL);

        System.out.printf("导出完成：总发票 %d（已排除开票时间早于2024-06-18的 %d 张）, sheet1=%d, sheet2=%d, sheet3=%d, sheet4组明细行=%d, sheet5=%d, sheet6=%d%n",
                invoices.size(), beforeDateCount, sheet1.size(), sheet2.size(), sheet3.size(), sheet4Rows.size(), sheet5.size(), sheet6.size());
        System.out.println("输出文件：" + OUTPUT_EXCEL.toAbsolutePath());
    }

    private static List<Path> collectXmlFiles(Path dir) throws Exception {
        if (!Files.exists(dir) || !Files.isDirectory(dir)) {
            return List.of();
        }
        try (Stream<Path> stream = Files.walk(dir)) {
            return stream.filter(p -> Files.isRegularFile(p)
                            && p.getFileName().toString().toLowerCase().endsWith(".xml"))
                    .sorted()
                    .collect(Collectors.toList());
        }
    }

    private static List<InvoiceRecord> parseAll(List<Path> xmlFiles) {
        List<InvoiceRecord> result = new ArrayList<>();
        for (Path xml : xmlFiles) {
            try {
                InvoiceRecord record = parseOne(xml);
                if (record != null) {
                    result.add(record);
                }
            } catch (Exception ex) {
                System.out.printf("[解析失败] %s | %s%n", xml.getFileName(), ex.getMessage());
            }
        }
        return result;
    }

    private static InvoiceRecord parseOne(Path xmlPath) throws Exception {
        try (InputStream in = Files.newInputStream(xmlPath)) {
            Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            doc.getDocumentElement().normalize();

            String taxBureauName = textOfFirst(doc, "TaxBureauName");
            String invoiceNumber = textOfFirst(doc, "InvoiceNumber");
            if (invoiceNumber.isBlank()) {
                invoiceNumber = textOfFirst(doc, "EIid");
            }
            String issueTime = textOfFirst(doc, "IssueTime");
            String sellerIdNum = textOfFirst(doc, "SellerIdNum");
            String sellerName = textOfFirst(doc, "SellerName");
            BigDecimal totalAmWithoutTax = parseDecimal(textOfFirst(doc, "TotalAmWithoutTax"));

            List<String> items = new ArrayList<>();
            NodeList itemNodes = doc.getElementsByTagName("IssuItemInformation");
            for (int i = 0; i < itemNodes.getLength(); i++) {
                Element el = (Element) itemNodes.item(i);
                String itemName = childText(el, "ItemName");
                if (!itemName.isBlank()) {
                    items.add(itemName.trim());
                }
            }
            if (items.isEmpty()) {
                items.add("");
            }
            items = dedupeItems(items);

            return new InvoiceRecord(
                    taxBureauName.trim(),
                    invoiceNumber.trim(),
                    issueTime.trim(),
                    sellerIdNum.trim(),
                    sellerName.trim(),
                    items,
                    totalAmWithoutTax,
                    xmlPath
            );
        }
    }

    private static List<GroupRow> buildSheet4Groups(List<InvoiceRecord> invoices) {
        Map<String, List<InvoiceRecord>> bySeller = invoices.stream()
                .collect(Collectors.groupingBy(InvoiceRecord::sellerIdNum, LinkedHashMap::new, Collectors.toList()));
        List<String> sortedSellerIds = new ArrayList<>(bySeller.keySet());
        sortedSellerIds.sort(Comparator.nullsLast(String::compareTo));

        List<GroupRow> rows = new ArrayList<>();
        int groupNo = 1;

        for (String sellerId : sortedSellerIds) {
            List<InvoiceRecord> list = new ArrayList<>(bySeller.getOrDefault(sellerId, Collections.emptyList()));
            list.sort(Comparator.comparing(InvoiceRecord::totalAmWithoutTax)
                    .thenComparing(InvoiceFolderWorkbookApp::safeIssueTime));

            List<GroupRow> sellerRows = new ArrayList<>();
            List<InvoiceRecord> current = new ArrayList<>();
            BigDecimal sum = BigDecimal.ZERO;
            for (InvoiceRecord inv : list) {
                BigDecimal amount = inv.totalAmWithoutTax();
                if (amount.compareTo(BigDecimal.ZERO) <= 0) {
                    // 负数/零发票直接标记为零散
                    rows.add(GroupRow.of(groupNo++, "零散不足100(负/零)", List.of(inv), amount));
                    continue;
                }

                current.add(inv);
                sum = sum.add(amount);
                if (sum.compareTo(new BigDecimal("100")) >= 0) {
                    sellerRows.add(GroupRow.of(0, "可凑满100+", current, sum));
                    current = new ArrayList<>();
                    sum = BigDecimal.ZERO;
                }
            }

            if (!current.isEmpty()) {
                if (!sellerRows.isEmpty()) {
                    // 同 SellerIdNum 已有可并入分组时，把不足100尾单并入最近一个分组（最后一个）
                    GroupRow last = sellerRows.get(sellerRows.size() - 1);
                    List<InvoiceRecord> mergedInvoices = new ArrayList<>(last.invoices());
                    mergedInvoices.addAll(current);
                    BigDecimal mergedSum = last.groupSum().add(sum);
                    sellerRows.set(sellerRows.size() - 1, GroupRow.of(0, last.groupTag(), mergedInvoices, mergedSum));
                } else {
                    // 没有同 SellerIdNum 可并入分组，且金额不足100，才保留零散不足100
                    sellerRows.add(GroupRow.of(0, "零散不足100", current, sum));
                }
            }

            for (GroupRow sellerRow : sellerRows) {
                rows.add(GroupRow.of(groupNo++, sellerRow.groupTag(), sellerRow.invoices(), sellerRow.groupSum()));
            }
        }

        return rows;
    }

    private static void writeWorkbook(List<InvoiceRecord> sheet1,
                                      List<InvoiceRecord> sheet2,
                                      List<InvoiceRecord> sheet3,
                                      List<GroupRow> sheet4Rows,
                                      List<InvoiceRecord> sheet5,
                                      List<InvoiceRecord> sheet6,
                                      Path output) throws Exception {
        Path parent = output.toAbsolutePath().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet s1 = wb.createSheet("sheet1_价税合计>100");
            Sheet s2 = wb.createSheet("sheet2_明细不在sheet1");
            Sheet s3 = wb.createSheet("sheet3_与sheet1冲突");
            Sheet s4 = wb.createSheet("sheet4_凑100分组");
            Sheet s5 = wb.createSheet("sheet5_sheet4排除列表命中");
            Sheet s6 = wb.createSheet("sheet6_零散不足100");

            Styles styles = buildStyles(wb);
            writeInvoiceSheet(s1, sheet1, BASE_HEADERS, styles);
            writeInvoiceSheet(s2, sheet2, BASE_HEADERS, styles);
            writeInvoiceSheet(s3, sheet3, BASE_HEADERS, styles);
            writeSheet4(s4, sheet4Rows, styles);
            writeInvoiceSheet(s5, sheet5, BASE_HEADERS, styles);
            writeInvoiceSheet(s6, sheet6, BASE_HEADERS, styles);

            try (OutputStream out = Files.newOutputStream(output)) {
                wb.write(out);
            }
        }
    }

    private static void writeInvoiceSheet(Sheet sheet, List<InvoiceRecord> invoices, String[] headers, Styles styles) {
        int rowCursor = 0;
        for (int i = 0; i < invoices.size(); i++) {
            if (i > 0) {
                rowCursor++;
            }
            InvoiceRecord inv = invoices.get(i);
            StylePair pair = styles.forIndex(i);
            rowCursor += writeInvoiceBlock(sheet, rowCursor, inv, headers, pair);
        }
        autoWidth(sheet, headers.length);
    }

    private static int writeInvoiceBlock(Sheet sheet, int headerRow, InvoiceRecord inv, String[] headers, StylePair pair) {
        Row h = sheet.createRow(headerRow);
        for (int c = 0; c < headers.length; c++) {
            Cell cell = h.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(pair.headerStyle());
        }

        int detailStart = headerRow + 1;
        int lines = inv.items().size();
        for (int i = 0; i < lines; i++) {
            Row row = sheet.createRow(detailStart + i);
            for (int c = 0; c < headers.length; c++) {
                row.createCell(c).setCellStyle(pair.bodyStyle());
            }
            row.getCell(5).setCellValue(inv.items().get(i));
        }
        int detailEnd = detailStart + lines - 1;

        writeMergedText(sheet, 0, detailStart, detailEnd, inv.taxBureauName(), pair.bodyStyle());
        writeMergedText(sheet, 1, detailStart, detailEnd, inv.invoiceNumber(), pair.bodyStyle());
        writeMergedText(sheet, 2, detailStart, detailEnd, inv.issueTime(), pair.bodyStyle());
        writeMergedText(sheet, 3, detailStart, detailEnd, inv.sellerIdNum(), pair.bodyStyle());
        writeMergedText(sheet, 4, detailStart, detailEnd, inv.sellerName(), pair.bodyStyle());
        writeMergedNumber(sheet, 6, detailStart, detailEnd, inv.totalAmWithoutTax(), pair.bodyStyle());
        return 1 + lines;
    }

    private static void writeSheet4(Sheet sheet, List<GroupRow> groupRows, Styles styles) {
        int rowCursor = 0;
        int blockIndex = 0;
        for (GroupRow group : groupRows) {
            if (rowCursor > 0) {
                rowCursor++;
            }
            StylePair pair = styles.forIndex(blockIndex++);

            Row h = sheet.createRow(rowCursor++);
            for (int c = 0; c < SHEET4_HEADERS.length; c++) {
                Cell cell = h.createCell(c);
                cell.setCellValue(SHEET4_HEADERS[c]);
                cell.setCellStyle(pair.headerStyle());
            }

            int detailStart = rowCursor;
            int detailLines = 0;
            List<DetailLine> detailLinesOrdered = buildDetailLinesSortedByIssueTime(group.invoices());
            for (DetailLine line : detailLinesOrdered) {
                InvoiceRecord inv = line.invoice();
                Row r = sheet.createRow(rowCursor++);
                detailLines++;
                for (int c = 0; c < SHEET4_HEADERS.length; c++) {
                    r.createCell(c).setCellStyle(pair.bodyStyle());
                }
                r.getCell(2).setCellValue(inv.taxBureauName());
                r.getCell(3).setCellValue(inv.invoiceNumber());
                r.getCell(4).setCellValue(inv.issueTime());
                r.getCell(5).setCellValue(inv.sellerIdNum());
                r.getCell(6).setCellValue(inv.sellerName());
                r.getCell(7).setCellValue(line.item());
                r.getCell(8).setCellValue(inv.totalAmWithoutTax().doubleValue());
            }
            int detailEnd = detailStart + Math.max(detailLines - 1, 0);
            writeMergedText(sheet, 0, detailStart, detailEnd, "G" + group.groupNo(), pair.bodyStyle());
            writeMergedText(sheet, 1, detailStart, detailEnd, group.groupTag(), pair.bodyStyle());
            writeMergedNumber(sheet, 9, detailStart, detailEnd, group.groupSum(), pair.bodyStyle());
        }
        autoWidth(sheet, SHEET4_HEADERS.length);
    }

    private static void writeMergedText(Sheet sheet, int col, int firstRow, int lastRow, String value, CellStyle style) {
        if (firstRow < lastRow) {
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(firstRow, lastRow, col, col));
        }
        Row row = sheet.getRow(firstRow);
        if (row == null) {
            row = sheet.createRow(firstRow);
        }
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        cell.setCellStyle(style);
        cell.setCellValue(value == null ? "" : value);
    }

    private static void writeMergedNumber(Sheet sheet, int col, int firstRow, int lastRow, BigDecimal value, CellStyle style) {
        if (firstRow < lastRow) {
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(firstRow, lastRow, col, col));
        }
        Row row = sheet.getRow(firstRow);
        if (row == null) {
            row = sheet.createRow(firstRow);
        }
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        cell.setCellStyle(style);
        cell.setCellValue(value.doubleValue());
    }

    private static void autoWidth(Sheet sheet, int columns) {
        int[] maxUnits = new int[columns];
        for (int i = 0; i < columns; i++) {
            maxUnits[i] = 10;
        }
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            for (int c = 0; c < columns; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) {
                    continue;
                }
                String txt = switch (cell.getCellType()) {
                    case STRING -> cell.getStringCellValue();
                    case NUMERIC -> String.valueOf(cell.getNumericCellValue());
                    case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                    default -> "";
                };
                if (!txt.isBlank()) {
                    maxUnits[c] = Math.max(maxUnits[c], displayUnits(txt));
                }
            }
        }
        for (int i = 0; i < columns; i++) {
            sheet.setColumnWidth(i, Math.min((maxUnits[i] + 2) * 256, 255 * 256));
        }
    }

    private static int displayUnits(String txt) {
        int units = 0;
        for (char c : txt.toCharArray()) {
            units += c > 127 ? 2 : 1;
        }
        return units;
    }

    private static Styles buildStyles(XSSFWorkbook wb) {
        List<StylePair> pairs = new ArrayList<>();
        for (int i = 0; i < BLOCK_HEADER_COLORS.length; i++) {
            CellStyle header = createStyle(wb, BLOCK_HEADER_COLORS[i], true);
            CellStyle body = createStyle(wb, BLOCK_BODY_COLORS[i], false);
            pairs.add(new StylePair(header, body));
        }
        return new Styles(pairs);
    }

    private static CellStyle createStyle(XSSFWorkbook wb, byte[] rgb, boolean bold) {
        CellStyle style = wb.createCellStyle();
        if (style instanceof XSSFCellStyle xssf) {
            xssf.setFillForegroundColor(new XSSFColor(rgb, null));
        }
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setWrapText(true);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        if (bold) {
            Font f = wb.createFont();
            f.setBold(true);
            style.setFont(f);
        }
        return style;
    }

    private static String textOfFirst(Document doc, String tagName) {
        NodeList list = doc.getElementsByTagName(tagName);
        if (list.getLength() == 0) {
            return "";
        }
        return list.item(0).getTextContent();
    }

    private static String childText(Element parent, String tagName) {
        NodeList nodes = parent.getElementsByTagName(tagName);
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            if (node.getParentNode() == parent) {
                return node.getTextContent();
            }
        }
        return "";
    }

    private static BigDecimal parseDecimal(String value) {
        try {
            return new BigDecimal(value.trim());
        } catch (Exception ex) {
            return BigDecimal.ZERO;
        }
    }

    private static List<DetailLine> buildDetailLinesSortedByIssueTime(List<InvoiceRecord> invoices) {
        List<DetailLine> lines = new ArrayList<>();
        Set<String> seen = new HashSet<>();
        for (InvoiceRecord inv : invoices) {
            for (String item : inv.items()) {
                String key = inv.invoiceNumber() + "|" + normItem(item);
                if (!seen.add(key)) {
                    continue;
                }
                lines.add(new DetailLine(inv, item));
            }
        }
        lines.sort(Comparator.comparing((DetailLine line) -> safeIssueTime(line.invoice()))
                .thenComparing(line -> line.invoice().invoiceNumber())
                .thenComparing(DetailLine::item));
        return lines;
    }

    private static List<String> dedupeItems(List<String> items) {
        List<String> result = new ArrayList<>();
        Set<String> seen = new HashSet<>();
        for (String item : items) {
            String key = normItem(item);
            if (seen.add(key)) {
                result.add(item);
            }
        }
        return result.isEmpty() ? List.of("") : result;
    }

    private static LocalDateTime safeIssueTime(InvoiceRecord inv) {
        return parseIssueTime(inv.issueTime());
    }

    private static LocalDateTime parseIssueTime(String issueTime) {
        if (issueTime == null || issueTime.isBlank()) {
            return LocalDateTime.MIN;
        }
        String value = issueTime.trim();
        try {
            return LocalDateTime.parse(value, ISSUE_TIME_FORMAT);
        } catch (DateTimeParseException ignored) {
            // try other common formats
        }
        try {
            return LocalDateTime.parse(value, DateTimeFormatter.ISO_LOCAL_DATE_TIME);
        } catch (DateTimeParseException ignored) {
            // continue
        }
        try {
            return LocalDate.parse(value, DateTimeFormatter.ISO_LOCAL_DATE).atStartOfDay();
        } catch (DateTimeParseException ignored) {
            return LocalDateTime.MIN;
        }
    }

    private record DetailLine(InvoiceRecord invoice, String item) {
    }

    private static String normItem(String item) {
        if (item == null) {
            return "";
        }
        return item.replaceAll("\\s+", "").trim();
    }

    private record InvoiceRecord(String taxBureauName,
                                 String invoiceNumber,
                                 String issueTime,
                                 String sellerIdNum,
                                 String sellerName,
                                 List<String> items,
                                 BigDecimal totalAmWithoutTax,
                                 Path sourcePath) {
    }

    private record StylePair(CellStyle headerStyle, CellStyle bodyStyle) {
    }

    private record Styles(List<StylePair> pairs) {
        StylePair forIndex(int index) {
            return pairs.get(index % pairs.size());
        }
    }

    private record GroupRow(int groupNo, String groupTag, List<InvoiceRecord> invoices, BigDecimal groupSum) {
        static GroupRow of(int groupNo, String groupTag, List<InvoiceRecord> invoices, BigDecimal sum) {
            return new GroupRow(groupNo, groupTag, new ArrayList<>(invoices), sum);
        }
    }
}
