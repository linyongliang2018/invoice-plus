package org.example.invoice;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.invoice.model.InvoiceData;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class InvoiceExcelExporter {

    private static final DateTimeFormatter ISSUE_TIME_FORMAT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    /** 交替底色：表头略深、明细略浅 */
    private static final byte[][] BLOCK_HEADER_COLORS = {
            {(byte) 255, (byte) 230, (byte) 153},
            {(byte) 198, (byte) 224, (byte) 180}
    };
    private static final byte[][] BLOCK_BODY_COLORS = {
            {(byte) 255, (byte) 242, (byte) 204},
            {(byte) 226, (byte) 239, (byte) 218}
    };

    private static final String[] HEADERS = {
            "发票机构(TaxBureauName)",
            "发票号码(EIid)",
            "开票纳税人识别号(SellerIdNum)",
            "开票公司(SellerName)",
            "开票商品明细(ItemName)",
            "明细价税合计(TotalTaxIncludedAmount)",
            "开票总金额价税合计(TotalTax-includedAmount)",
            "开票时间(IssueTime)"
    };

    /**
     * 全量覆盖写入：按开票时间升序，时间相同按发票号码升序；列宽按内容自动调整。
     */
    public ExportResult exportInvoices(Path excelPath, List<InvoiceData> invoices) throws IOException {
        List<InvoiceData> prepared = prepareInvoices(invoices);

        Path parent = excelPath.toAbsolutePath().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("发票汇总");
            BlockStyles blockStyles = createBlockStyles(workbook);
            int currentRow = 0;
            for (int i = 0; i < prepared.size(); i++) {
                if (i > 0) {
                    currentRow++;
                }
                currentRow += writeInvoiceBlock(sheet, currentRow, prepared.get(i), blockStyles.forIndex(i));
            }
            adjustColumnWidths(sheet);

            try (OutputStream out = Files.newOutputStream(excelPath)) {
                workbook.write(out);
            }
        }

        int skipped = invoices.size() - prepared.size();
        return new ExportResult(prepared.size(), skipped);
    }

    private static List<InvoiceData> prepareInvoices(List<InvoiceData> invoices) {
        Map<String, InvoiceData> unique = new LinkedHashMap<>();
        for (InvoiceData invoice : invoices) {
            String eiid = safe(invoice.getEiid());
            if (eiid.isEmpty()) {
                continue;
            }
            unique.putIfAbsent(eiid, invoice);
        }

        List<InvoiceData> result = new ArrayList<>(unique.values());
        result.sort(Comparator
                .comparing(InvoiceExcelExporter::parseIssueTime)
                .thenComparing(invoice -> safe(invoice.getEiid())));
        return result;
    }

    private static LocalDateTime parseIssueTime(InvoiceData invoice) {
        String issueTime = safe(invoice.getIssueTime());
        if (issueTime.isEmpty()) {
            return LocalDateTime.MIN;
        }
        try {
            return LocalDateTime.parse(issueTime, ISSUE_TIME_FORMAT);
        } catch (DateTimeParseException ex) {
            return LocalDateTime.MIN;
        }
    }

    /** @return 本发票块占用的行数（表头 + 明细） */
    private static int writeInvoiceBlock(Sheet sheet, int headerRowIndex, InvoiceData invoice,
                                         BlockStylePair styles) {
        Row headerRow = sheet.createRow(headerRowIndex);
        for (int col = 0; col < HEADERS.length; col++) {
            Cell cell = headerRow.createCell(col);
            cell.setCellValue(HEADERS[col]);
            cell.setCellStyle(styles.header());
        }

        int detailStartRow = headerRowIndex + 1;
        List<InvoiceData.LineItem> items = invoice.getLineItems();
        int detailCount = items.size();

        for (int i = 0; i < detailCount; i++) {
            Row row = sheet.createRow(detailStartRow + i);
            for (int col = 0; col < HEADERS.length; col++) {
                row.createCell(col).setCellStyle(styles.body());
            }
            InvoiceData.LineItem item = items.get(i);
            row.getCell(4).setCellValue(safe(item.getItemName()));
            setNumericOrText(row.getCell(5), item.getTotalTaxIncludedAmount());
        }

        int detailEndRow = detailStartRow + Math.max(detailCount - 1, 0);
        if (detailCount > 0) {
            mergeTextColumn(sheet, 0, detailStartRow, detailEndRow, safe(invoice.getTaxBureauName()), styles.body());
            mergeTextColumn(sheet, 1, detailStartRow, detailEndRow, safe(invoice.getEiid()), styles.body());
            mergeTextColumn(sheet, 2, detailStartRow, detailEndRow, safe(invoice.getSellerIdNum()), styles.body());
            mergeTextColumn(sheet, 3, detailStartRow, detailEndRow, safe(invoice.getSellerName()), styles.body());
            mergeAmountColumn(sheet, 6, detailStartRow, detailEndRow, safe(invoice.getTotalTaxIncludedAmount()), styles.body());
            mergeTextColumn(sheet, 7, detailStartRow, detailEndRow, safe(invoice.getIssueTime()), styles.body());
        }

        return 1 + detailCount;
    }

    private static BlockStyles createBlockStyles(Workbook workbook) {
        CellStyle[] headers = new CellStyle[BLOCK_HEADER_COLORS.length];
        CellStyle[] bodies = new CellStyle[BLOCK_BODY_COLORS.length];
        for (int i = 0; i < BLOCK_HEADER_COLORS.length; i++) {
            headers[i] = createFillStyle(workbook, BLOCK_HEADER_COLORS[i], true);
            bodies[i] = createFillStyle(workbook, BLOCK_BODY_COLORS[i], false);
        }
        return new BlockStyles(headers, bodies);
    }

    private static CellStyle createFillStyle(Workbook workbook, byte[] rgb, boolean bold) {
        CellStyle style = workbook.createCellStyle();
        if (style instanceof XSSFCellStyle xssfStyle) {
            xssfStyle.setFillForegroundColor(new XSSFColor(rgb, null));
        }
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        if (bold) {
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
        }
        return style;
    }

    private record BlockStylePair(CellStyle header, CellStyle body) {
    }

    private record BlockStyles(CellStyle[] headers, CellStyle[] bodies) {
        BlockStylePair forIndex(int invoiceIndex) {
            int i = invoiceIndex % headers.length;
            return new BlockStylePair(headers[i], bodies[i]);
        }
    }

    private static void adjustColumnWidths(Sheet sheet) {
        int[] maxUnits = new int[HEADERS.length];
        for (int col = 0; col < HEADERS.length; col++) {
            maxUnits[col] = displayWidthUnits(HEADERS[col]);
        }

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            for (int col = 0; col < HEADERS.length; col++) {
                Cell cell = row.getCell(col);
                if (cell == null) {
                    continue;
                }
                String text = cellToString(cell);
                if (!text.isEmpty()) {
                    maxUnits[col] = Math.max(maxUnits[col], displayWidthUnits(text));
                }
            }
        }

        for (int col = 0; col < HEADERS.length; col++) {
            int width = Math.min((maxUnits[col] + 2) * 256, 255 * 256);
            sheet.setColumnWidth(col, width);
        }
    }

    /** 中文等宽字符按 2 单位、英文按 1 单位估算显示宽度 */
    private static int displayWidthUnits(String text) {
        int units = 0;
        for (char c : text.toCharArray()) {
            units += c > 127 ? 2 : 1;
        }
        return units;
    }

    private static String cellToString(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                double value = cell.getNumericCellValue();
                if (value == Math.rint(value) && !Double.isInfinite(value)) {
                    yield String.valueOf((long) value);
                }
                yield String.valueOf(value);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    private static void mergeTextColumn(Sheet sheet, int col, int firstRow, int lastRow,
                                        String value, CellStyle style) {
        if (firstRow < lastRow) {
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, col, col));
        }
        Row row = sheet.getRow(firstRow);
        if (row == null) {
            row = sheet.createRow(firstRow);
        }
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static void mergeAmountColumn(Sheet sheet, int col, int firstRow, int lastRow,
                                          String value, CellStyle style) {
        if (firstRow < lastRow) {
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, col, col));
        }
        Row row = sheet.getRow(firstRow);
        if (row == null) {
            row = sheet.createRow(firstRow);
        }
        Cell cell = row.getCell(col);
        if (cell == null) {
            cell = row.createCell(col);
        }
        setNumericOrText(cell, value);
        cell.setCellStyle(style);
    }

    private static void setNumericOrText(Cell cell, String value) {
        if (value == null || value.isBlank()) {
            cell.setBlank();
            return;
        }
        try {
            cell.setCellValue(Double.parseDouble(value));
        } catch (NumberFormatException ex) {
            cell.setCellValue(value);
        }
    }

    private static String safe(String value) {
        return value == null ? "" : value.trim();
    }

    public record ExportResult(int exported, int skipped) {
    }
}
