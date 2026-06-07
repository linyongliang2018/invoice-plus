package org.example.invoice;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * 读取已整理的多 sheet 报表（sheet4 指定分组 + sheet6 + 额外发票号），
 * 再从 XML 重新解析，按 SellerIdNum 使用价税合计凑100分组。
 * <p>
 * 运行前修改 {@link #PACK_MODE} 选择算法即可。
 */
public class InvoiceOptimalRepackApp {

    private static final Path INPUT_EXCEL = Path.of("D:\\发票整理-多sheet.xlsx");
    private static final Path XML_DIR = Path.of("C:\\Users\\40696\\Desktop\\京东发票-京东发票明细");
    private static final Path OUTPUT_EXCEL = Path.of("D:\\发票二次凑100.xlsx");

    private static final String SHEET4_NAME = "sheet4_凑100分组";
    private static final String SHEET6_NAME = "sheet6_零散不足100";
    private static final BigDecimal TARGET = new BigDecimal("100");
    /** 折半枚举适用上限：n=26 时约 6700 万次比较/组，仍可在秒级完成 */
    private static final int MITM_MAX_SIZE = 26;
    /** 超过此数量时提示最优算法可能较慢 */
    private static final int OPTIMAL_SLOW_WARN_INVOICES = 20;

    private enum PackMode {
        /** seed+grow 贪心，最快 */
        GREEDY("贪心"),
        /** 每步最小浪费 + 折半枚举，速度与质量折中 */
        HEURISTIC("启发"),
        /** 位掩码 DP，分组数最多，发票多时可能较慢 */
        OPTIMAL("最优");

        private final String label;

        PackMode(String label) {
            this.label = label;
        }

        String label() {
            return label;
        }
    }

    /** 凑组算法：改枚举值后直接在 IDEA 运行 main */
    private static final PackMode PACK_MODE = PackMode.OPTIMAL;

    private static final DateTimeFormatter ISSUE_TIME_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    private static final Set<String> TARGET_GROUP_NOS = Set.of(
            "G8", "G24", "G26", "G27", "G29", "G30", "G31"
    );

    private static final List<String> EXTRA_INVOICE_NUMBERS = List.of(
            "26447000001040855280",
            "26447000001117493196",
            "26467000000075151398",
            "26957000000105472001",
            "26467000000050876172",
            "26957000000082848580",
            "26957000000036136557",
            "25117000001554546781",
            "25517000000542707055",
            "25117000001228254171",
            "25117000001226721471",
            "25117000001225338658",
            "25957000000088242009",
            "25447000000631134692",
            "25957000000059922559",
            "25447000000367165132",
            "25447200000118166264",
            "24447000000494097861",
            "24447000000443274858",
            "24447000000439301686",
            "24427000000102259566",
            "24467000000009392562",
            "26447000000747832213",
            "26447000000767698796",
            "25447000000547589094",
            "26447000000672611747",
            "26447000000672611837",
            "25447000001217254144",
            "26447000000646353802",
            "26447000001115078946",
            "26447000000671312988",
            "26447000000726600609",
            "25447000001506810266",
            "25447200000118463217",
            "25447000001211248188",
            "26217000000246686557",
            "26117000000805575574",
            "26447000001142096297",
            "26447000001142096326",
            "26117000000805575592",
            "26117000000805575584",
            "26117000000805575578",
            "26117000000813196883",
            "26117000000808206378"
    );

    private static final String[] GROUP_HEADERS = {
            "分组编号(GroupNo)",
            "分组标签(GroupTag)",
            "开票机构(TaxBureauName)",
            "发票号码(InvoiceNumber)",
            "开票时间(IssueTime)",
            "销方纳税人识别号(SellerIdNum)",
            "销方名称(SellerName)",
            "商品明细(ItemName)",
            "价税合计(TotalTax-includedAmount)",
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
        System.out.println("凑组算法：" + PACK_MODE.label() + " (" + PACK_MODE + ")");

        Set<String> scopeNumbers = collectScopeInvoiceNumbers(INPUT_EXCEL);
        scopeNumbers.addAll(EXTRA_INVOICE_NUMBERS);
        System.out.printf("范围发票数：%d%n", scopeNumbers.size());

        List<InvoiceRecord> scopedInvoices = parseScopedInvoices(XML_DIR, scopeNumbers);
        if (scopedInvoices.isEmpty()) {
            System.out.println("范围内未匹配到任何 XML 发票。");
            return;
        }

        Set<String> matched = scopedInvoices.stream()
                .map(InvoiceRecord::invoiceNumber)
                .collect(Collectors.toSet());
        Set<String> missing = new LinkedHashSet<>(scopeNumbers);
        missing.removeAll(matched);
        if (!missing.isEmpty()) {
            System.out.printf("警告：范围中有 %d 张发票未在 XML 目录找到。%n", missing.size());
        }

        Map<String, List<InvoiceRecord>> bySeller = scopedInvoices.stream()
                .collect(Collectors.groupingBy(InvoiceRecord::sellerIdNum, LinkedHashMap::new, Collectors.toList()));

        List<String> sellerIds = new ArrayList<>(bySeller.keySet());
        sellerIds.sort(Comparator.nullsLast(String::compareTo));
        System.out.printf("解析完成：%d 张发票，分布在 %d 个销方，开始并行凑组…%n",
                scopedInvoices.size(), sellerIds.size());

        List<SellerPackResult> sellerResults = sellerIds.parallelStream()
                .map(sellerId -> packOneSeller(sellerId, bySeller.get(sellerId), PACK_MODE))
                .sorted(Comparator.comparing(SellerPackResult::sellerId, Comparator.nullsLast(String::compareTo)))
                .toList();

        List<GroupRow> groupedRows = new ArrayList<>();
        List<GroupRow> insufficientRows = new ArrayList<>();
        int nextGroupNo = 1;
        int insufficientGroupNo = 1;

        for (SellerPackResult sellerResult : sellerResults) {
            String sellerId = sellerResult.sellerId();
            PackResult packResult = sellerResult.packResult();
            for (List<InvoiceRecord> group : packResult.groups()) {
                BigDecimal sum = group.stream()
                        .map(InvoiceRecord::totalTaxIncludedAmount)
                        .reduce(BigDecimal.ZERO, BigDecimal::add);
                groupedRows.add(new GroupRow(nextGroupNo++, "凑满100+", group, sum, sellerId));
            }
            if (!packResult.leftovers().isEmpty()) {
                List<InvoiceRecord> leftovers = new ArrayList<>(packResult.leftovers());
                leftovers.sort(Comparator.comparing(InvoiceOptimalRepackApp::safeIssueTime)
                        .thenComparing(InvoiceRecord::invoiceNumber));
                BigDecimal sum = leftovers.stream()
                        .map(InvoiceRecord::totalTaxIncludedAmount)
                        .reduce(BigDecimal.ZERO, BigDecimal::add);
                insufficientRows.add(new GroupRow(insufficientGroupNo++, "不足100", leftovers, sum, sellerId));
            }
        }

        int insufficientInvoiceCount = insufficientRows.stream()
                .mapToInt(row -> row.invoices().size())
                .sum();

        writeWorkbook(groupedRows, insufficientRows, OUTPUT_EXCEL);

        System.out.printf("导出完成：sheet1 可凑满100+ 分组 %d 个，sheet2 不足100 分组 %d 个（共 %d 张发票）。%n",
                groupedRows.size(), insufficientRows.size(), insufficientInvoiceCount);
        System.out.println("输出文件：" + OUTPUT_EXCEL.toAbsolutePath());
    }

    private static Set<String> collectScopeInvoiceNumbers(Path excelPath) throws Exception {
        Set<String> scope = new LinkedHashSet<>();
        try (InputStream in = Files.newInputStream(excelPath);
             var workbook = WorkbookFactory.create(in)) {
            scope.addAll(readSheet4InvoiceNumbers(workbook.getSheet(SHEET4_NAME)));
            scope.addAll(readSheet6InvoiceNumbers(workbook.getSheet(SHEET6_NAME)));
        }
        return scope;
    }

    private static Set<String> readSheet4InvoiceNumbers(Sheet sheet) {
        Set<String> numbers = new LinkedHashSet<>();
        if (sheet == null) {
            System.out.println("未找到 sheet4：" + SHEET4_NAME);
            return numbers;
        }

        String currentGroup = "";
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String groupNo = cellString(row.getCell(0));
            if (groupNo.startsWith("G")) {
                currentGroup = groupNo.trim();
            }
            if (!TARGET_GROUP_NOS.contains(currentGroup)) {
                continue;
            }
            String invoiceNumber = cellString(row.getCell(3));
            if (isInvoiceNumber(invoiceNumber)) {
                numbers.add(invoiceNumber);
            }
        }
        return numbers;
    }

    private static Set<String> readSheet6InvoiceNumbers(Sheet sheet) {
        Set<String> numbers = new LinkedHashSet<>();
        if (sheet == null) {
            System.out.println("未找到 sheet6：" + SHEET6_NAME);
            return numbers;
        }

        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String col0 = cellString(row.getCell(0));
            if (col0.contains("TaxBureauName") || col0.contains("开票机构")) {
                continue;
            }
            String invoiceNumber = cellString(row.getCell(1));
            if (isInvoiceNumber(invoiceNumber)) {
                numbers.add(invoiceNumber);
            }
        }
        return numbers;
    }

    private static SellerPackResult packOneSeller(String sellerId, List<InvoiceRecord> sellerInvoices, PackMode mode) {
        int count = sellerInvoices == null ? 0 : sellerInvoices.size();
        System.out.printf("[开始] 销方 %s，%d 张发票，算法 %s（线程 %s）%n",
                sellerId, count, mode.label(), Thread.currentThread().getName());
        long start = System.currentTimeMillis();
        PackResult packResult = packForSeller(sellerInvoices, mode);
        System.out.printf("[完成] 销方 %s，耗时 %d ms，可凑分组 %d，不足100 %d 张%n",
                sellerId,
                System.currentTimeMillis() - start,
                packResult.groups().size(),
                packResult.leftovers().size());
        return new SellerPackResult(sellerId, packResult);
    }

    private static List<InvoiceRecord> parseScopedInvoices(Path xmlDir, Set<String> scopeNumbers) throws Exception {
        List<Path> xmlFiles = collectXmlFiles(xmlDir);
        Map<String, InvoiceRecord> unique = new ConcurrentHashMap<>();

        xmlFiles.parallelStream().forEach(xml -> {
            try {
                InvoiceRecord record = parseOne(xml);
                if (record == null || record.invoiceNumber().isBlank()) {
                    return;
                }
                if (!scopeNumbers.contains(record.invoiceNumber())) {
                    return;
                }
                unique.putIfAbsent(record.invoiceNumber(), record);
            } catch (Exception ex) {
                synchronized (System.out) {
                    System.out.printf("[解析失败] %s | %s%n", xml.getFileName(), ex.getMessage());
                }
            }
        });
        return new ArrayList<>(unique.values());
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

    private static InvoiceRecord parseOne(Path xmlPath) throws Exception {
        try (InputStream in = Files.newInputStream(xmlPath)) {
            Document doc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(in);
            doc.getDocumentElement().normalize();

            String invoiceNumber = textOfFirst(doc, "InvoiceNumber");
            if (invoiceNumber.isBlank()) {
                invoiceNumber = textOfFirst(doc, "EIid");
            }

            BigDecimal totalTaxIncludedAmount = parseDecimal(textOfFirst(doc, "TotalTax-includedAmount"));
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
                    textOfFirst(doc, "TaxBureauName").trim(),
                    invoiceNumber.trim(),
                    textOfFirst(doc, "IssueTime").trim(),
                    textOfFirst(doc, "SellerIdNum").trim(),
                    textOfFirst(doc, "SellerName").trim(),
                    items,
                    totalTaxIncludedAmount,
                    xmlPath
            );
        }
    }

    private static PackResult packForSeller(List<InvoiceRecord> invoices, PackMode mode) {
        List<InvoiceRecord> positive = new ArrayList<>();
        List<InvoiceRecord> leftovers = new ArrayList<>();
        for (InvoiceRecord inv : invoices) {
            if (inv.totalTaxIncludedAmount().compareTo(BigDecimal.ZERO) <= 0) {
                leftovers.add(inv);
            } else {
                positive.add(inv);
            }
        }
        if (positive.isEmpty()) {
            return new PackResult(List.of(), leftovers);
        }

        List<List<InvoiceRecord>> singleGroups = new ArrayList<>();
        List<InvoiceRecord> needPack = new ArrayList<>();
        for (InvoiceRecord inv : positive) {
            if (inv.totalTaxIncludedAmount().compareTo(TARGET) >= 0) {
                singleGroups.add(List.of(inv));
            } else {
                needPack.add(inv);
            }
        }

        if (needPack.isEmpty()) {
            return new PackResult(singleGroups, leftovers);
        }

        PackResult packed = switch (mode) {
            case GREEDY -> packGreedyIterative(needPack, leftovers);
            case HEURISTIC -> packMinWasteIterative(needPack, leftovers);
            case OPTIMAL -> packExactPositiveOnly(needPack, leftovers);
        };
        List<List<InvoiceRecord>> allGroups = new ArrayList<>(singleGroups);
        allGroups.addAll(packed.groups());
        return new PackResult(allGroups, packed.leftovers());
    }

    /** 贪心：每轮 seed+grow，选浪费最小的可行组 */
    private static PackResult packGreedyIterative(List<InvoiceRecord> needPack, List<InvoiceRecord> initialLeftovers) {
        List<InvoiceRecord> remaining = new ArrayList<>(needPack);
        List<List<InvoiceRecord>> groups = new ArrayList<>();
        List<InvoiceRecord> leftovers = new ArrayList<>(initialLeftovers);

        while (!remaining.isEmpty()) {
            List<InvoiceRecord> group = findBestGroupGreedy(remaining);
            if (group.isEmpty()) {
                leftovers.addAll(remaining);
                break;
            }
            groups.add(group);
            Set<String> groupedNumbers = group.stream()
                    .map(InvoiceRecord::invoiceNumber)
                    .collect(Collectors.toSet());
            remaining.removeIf(inv -> groupedNumbers.contains(inv.invoiceNumber()));
        }
        return new PackResult(groups, leftovers);
    }

    /**
     * 精确最优：位掩码 DP + 记忆化，最大化「和 >= 100」的分组数量。
     */
    private static PackResult packExactPositiveOnly(List<InvoiceRecord> needPack, List<InvoiceRecord> initialLeftovers) {
        if (needPack.size() > OPTIMAL_SLOW_WARN_INVOICES) {
            System.out.printf("  提示：销方 %s 待组合 %d 张，最优算法可能较慢…%n",
                    needPack.get(0).sellerIdNum(), needPack.size());
        }
        if (needPack.size() >= Long.SIZE) {
            throw new IllegalStateException("单个销方待组合发票超过 " + (Long.SIZE - 1)
                    + " 张，最优算法不支持，请改用 heuristic 或 greedy。");
        }
        return packExact(needPack, initialLeftovers);
    }

    private static PackResult packExact(List<InvoiceRecord> positive, List<InvoiceRecord> initialLeftovers) {
        int n = positive.size();
        long[] cents = new long[n];
        long targetCents = toCents(TARGET);
        for (int i = 0; i < n; i++) {
            cents[i] = toCents(positive.get(i).totalTaxIncludedAmount());
        }

        long fullMask = (1L << n) - 1;
        Map<Long, Integer> memo = new HashMap<>();
        computeMaxGroups(0L, n, fullMask, cents, targetCents, memo);

        List<List<InvoiceRecord>> groups = new ArrayList<>();
        List<InvoiceRecord> leftovers = new ArrayList<>(initialLeftovers);
        long decidedMask = 0L;
        long groupedMask = 0L;
        while (decidedMask != fullMask) {
            int first = firstUnsetBit(decidedMask, n);
            long bestSubset = findBestSubset(first, decidedMask, n, fullMask, cents, targetCents, memo);
            if (bestSubset == 0L) {
                decidedMask |= 1L << first;
                continue;
            }
            List<InvoiceRecord> group = new ArrayList<>();
            for (int i = 0; i < n; i++) {
                if ((bestSubset & (1L << i)) != 0L) {
                    group.add(positive.get(i));
                }
            }
            groups.add(group);
            decidedMask |= bestSubset;
            groupedMask |= bestSubset;
        }
        for (int i = 0; i < n; i++) {
            if ((groupedMask & (1L << i)) == 0L) {
                leftovers.add(positive.get(i));
            }
        }
        return new PackResult(groups, leftovers);
    }

    private static int computeMaxGroups(long mask,
                                        int n,
                                        long fullMask,
                                        long[] cents,
                                        long targetCents,
                                        Map<Long, Integer> memo) {
        if (mask == fullMask) {
            return 0;
        }
        Integer cached = memo.get(mask);
        if (cached != null) {
            return cached;
        }

        int first = firstUnsetBit(mask, n);
        int best = computeMaxGroups(mask | (1L << first), n, fullMask, cents, targetCents, memo);

        long groupMask = 1L << first;
        long sum = cents[first];
        best = Math.max(best, searchGroupFrom(first, mask, groupMask, sum, n, fullMask, cents, targetCents, memo));

        memo.put(mask, best);
        return best;
    }

    private static int searchGroupFrom(int first,
                                       long mask,
                                       long groupMask,
                                       long sum,
                                       int n,
                                       long fullMask,
                                       long[] cents,
                                       long targetCents,
                                       Map<Long, Integer> memo) {
        if (sum >= targetCents) {
            return 1 + computeMaxGroups(mask | groupMask, n, fullMask, cents, targetCents, memo);
        }

        int best = 0;
        long available = (fullMask ^ mask) & ~groupMask;
        for (int i = first + 1; i < n; i++) {
            if ((available & (1L << i)) == 0L) {
                continue;
            }
            long nextSum = sum + cents[i];
            if (nextSum + maxRemainingSum(available & ~((1L << i) | groupMask), cents) < targetCents) {
                continue;
            }
            best = Math.max(best, searchGroupFrom(first, mask, groupMask | (1L << i), nextSum,
                    n, fullMask, cents, targetCents, memo));
        }
        return best;
    }

    private static long findBestSubset(int first,
                                       long mask,
                                       int n,
                                       long fullMask,
                                       long[] cents,
                                       long targetCents,
                                       Map<Long, Integer> memo) {
        int target = memo.getOrDefault(mask, 0);
        int skipScore = memo.getOrDefault(mask | (1L << first), 0);
        if (target <= skipScore) {
            return 0L;
        }

        SubsetHolder holder = new SubsetHolder();
        searchBestSubset(first, mask, 1L << first, cents[first], n, fullMask, cents, targetCents, memo, target, holder);
        return holder.bestMask;
    }

    private static void searchBestSubset(int first,
                                         long mask,
                                         long groupMask,
                                         long sum,
                                         int n,
                                         long fullMask,
                                         long[] cents,
                                         long targetCents,
                                         Map<Long, Integer> memo,
                                         int targetScore,
                                         SubsetHolder holder) {
        if (sum >= targetCents) {
            int score = 1 + memo.getOrDefault(mask | groupMask, 0);
            if (score == targetScore) {
                long waste = sum - targetCents;
                if (waste < holder.bestWaste) {
                    holder.bestWaste = waste;
                    holder.bestMask = groupMask;
                }
            }
            return;
        }

        long available = (fullMask ^ mask) & ~groupMask;
        for (int i = first + 1; i < n; i++) {
            if ((available & (1L << i)) == 0L) {
                continue;
            }
            long nextSum = sum + cents[i];
            if (nextSum + maxRemainingSum(available & ~((1L << i) | groupMask), cents) < targetCents) {
                continue;
            }
            searchBestSubset(first, mask, groupMask | (1L << i), nextSum, n, fullMask, cents,
                    targetCents, memo, targetScore, holder);
        }
    }

    private static long maxRemainingSum(long availableMask, long[] cents) {
        long sum = 0;
        for (int i = 0; i < cents.length; i++) {
            if ((availableMask & (1L << i)) != 0L) {
                sum += cents[i];
            }
        }
        return sum;
    }

    private static int firstUnsetBit(long mask, int n) {
        for (int i = 0; i < n; i++) {
            if ((mask & (1L << i)) == 0L) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 启发式：每轮从 remaining 中选出一步最优子集（最小浪费 >=100），移除后继续。
     */
    private static PackResult packMinWasteIterative(List<InvoiceRecord> needPack, List<InvoiceRecord> initialLeftovers) {
        List<InvoiceRecord> remaining = new ArrayList<>(needPack);
        List<List<InvoiceRecord>> groups = new ArrayList<>();
        List<InvoiceRecord> leftovers = new ArrayList<>(initialLeftovers);

        while (!remaining.isEmpty()) {
            List<InvoiceRecord> group = findBestGroupMinWaste(remaining);
            if (group.isEmpty()) {
                leftovers.addAll(remaining);
                break;
            }
            groups.add(group);
            Set<String> groupedNumbers = group.stream()
                    .map(InvoiceRecord::invoiceNumber)
                    .collect(Collectors.toSet());
            remaining.removeIf(inv -> groupedNumbers.contains(inv.invoiceNumber()));
        }
        return new PackResult(groups, leftovers);
    }

    /** 在 remaining 中找 sum>=100 且浪费最小的子集；n 不大时用折半枚举搜全部子集，否则退回贪心 */
    private static List<InvoiceRecord> findBestGroupMinWaste(List<InvoiceRecord> remaining) {
        if (remaining.size() <= MITM_MAX_SIZE) {
            return findBestGroupMeetInMiddle(remaining);
        }
        return findBestGroupGreedy(remaining);
    }

    /**
     * Meet-in-the-Middle：枚举左/右半段子集的和，合并求「>=100 且浪费最小」的组合。
     * 对当前 remaining 规模，这一步是精确最优，不是贪心近似。
     */
    private static List<InvoiceRecord> findBestGroupMeetInMiddle(List<InvoiceRecord> remaining) {
        int n = remaining.size();
        long[] cents = new long[n];
        for (int i = 0; i < n; i++) {
            cents[i] = toCents(remaining.get(i).totalTaxIncludedAmount());
        }
        long targetCents = toCents(TARGET);
        int mid = n / 2;

        List<HalfEntry> left = enumerateHalf(cents, 0, mid);
        List<HalfEntry> right = enumerateHalf(cents, mid, n);
        right.sort(Comparator.comparingLong(e -> e.sum));

        long bestWaste = Long.MAX_VALUE;
        int bestLeftMask = 0;
        int bestRightMask = 0;
        int bestSize = Integer.MAX_VALUE;

        for (HalfEntry le : left) {
            if (le.sum >= targetCents) {
                long waste = le.sum - targetCents;
                int size = Integer.bitCount(le.mask);
                if (waste < bestWaste || (waste == bestWaste && size < bestSize)) {
                    bestWaste = waste;
                    bestLeftMask = le.mask;
                    bestRightMask = 0;
                    bestSize = size;
                }
                continue;
            }
            long need = targetCents - le.sum;
            int idx = lowerBoundBySum(right, need);
            if (idx >= right.size()) {
                continue;
            }
            HalfEntry re = right.get(idx);
            long waste = le.sum + re.sum - targetCents;
            int size = Integer.bitCount(le.mask) + Integer.bitCount(re.mask);
            if (waste < bestWaste || (waste == bestWaste && size < bestSize)) {
                bestWaste = waste;
                bestLeftMask = le.mask;
                bestRightMask = re.mask;
                bestSize = size;
            }
        }

        if (bestWaste == Long.MAX_VALUE) {
            return List.of();
        }

        List<InvoiceRecord> group = new ArrayList<>();
        for (int i = 0; i < mid; i++) {
            if ((bestLeftMask & (1 << i)) != 0) {
                group.add(remaining.get(i));
            }
        }
        for (int i = mid; i < n; i++) {
            if ((bestRightMask & (1 << (i - mid))) != 0) {
                group.add(remaining.get(i));
            }
        }
        return group;
    }

    private static List<HalfEntry> enumerateHalf(long[] cents, int start, int end) {
        int k = end - start;
        int total = 1 << k;
        List<HalfEntry> list = new ArrayList<>(total);
        for (int mask = 0; mask < total; mask++) {
            long sum = 0;
            for (int b = 0; b < k; b++) {
                if ((mask & (1 << b)) != 0) {
                    sum += cents[start + b];
                }
            }
            list.add(new HalfEntry(sum, mask));
        }
        return list;
    }

    private static int lowerBoundBySum(List<HalfEntry> entries, long need) {
        int lo = 0;
        int hi = entries.size();
        while (lo < hi) {
            int m = (lo + hi) >>> 1;
            if (entries.get(m).sum < need) {
                lo = m + 1;
            } else {
                hi = m;
            }
        }
        return lo;
    }

    private static long toCents(BigDecimal amount) {
        return amount.movePointRight(2).setScale(0, RoundingMode.HALF_UP).longValueExact();
    }

    /** 发票数超过折半上限时的后备：seed + grow 贪心 */
    private static List<InvoiceRecord> findBestGroupGreedy(List<InvoiceRecord> remaining) {
        List<InvoiceRecord> bestGroup = List.of();
        BigDecimal bestWaste = null;

        List<InvoiceRecord> sorted = new ArrayList<>(remaining);
        sorted.sort(Comparator.comparing(InvoiceRecord::totalTaxIncludedAmount).reversed()
                .thenComparing(InvoiceOptimalRepackApp::safeIssueTime));

        for (InvoiceRecord seed : sorted) {
            List<InvoiceRecord> group = growGroupGreedy(seed, remaining);
            if (group.isEmpty()) {
                continue;
            }
            BigDecimal sum = sumAmounts(group);
            if (sum.compareTo(TARGET) < 0) {
                continue;
            }
            BigDecimal waste = sum.subtract(TARGET);
            if (bestWaste == null || waste.compareTo(bestWaste) < 0) {
                bestWaste = waste;
                bestGroup = group;
            }
        }
        return bestGroup;
    }

    private static List<InvoiceRecord> growGroupGreedy(InvoiceRecord seed, List<InvoiceRecord> remaining) {
        List<InvoiceRecord> group = new ArrayList<>();
        group.add(seed);
        BigDecimal sum = seed.totalTaxIncludedAmount();

        List<InvoiceRecord> pool = remaining.stream()
                .filter(inv -> !inv.invoiceNumber().equals(seed.invoiceNumber()))
                .sorted(Comparator.comparing(InvoiceRecord::totalTaxIncludedAmount).reversed()
                        .thenComparing(InvoiceOptimalRepackApp::safeIssueTime))
                .collect(Collectors.toCollection(ArrayList::new));

        while (sum.compareTo(TARGET) < 0 && !pool.isEmpty()) {
            int pickIdx = pickNextInvoice(pool, sum);
            if (pickIdx < 0) {
                break;
            }
            InvoiceRecord picked = pool.remove(pickIdx);
            group.add(picked);
            sum = sum.add(picked.totalTaxIncludedAmount());
        }
        return sum.compareTo(TARGET) >= 0 ? group : List.of();
    }

    /** 优先选能凑满且浪费最小的；否则选金额最大的推进进度 */
    private static int pickNextInvoice(List<InvoiceRecord> pool, BigDecimal currentSum) {
        int bestReachIdx = -1;
        BigDecimal bestWaste = null;
        for (int i = 0; i < pool.size(); i++) {
            BigDecimal next = currentSum.add(pool.get(i).totalTaxIncludedAmount());
            if (next.compareTo(TARGET) >= 0) {
                BigDecimal waste = next.subtract(TARGET);
                if (bestReachIdx < 0 || waste.compareTo(bestWaste) < 0) {
                    bestReachIdx = i;
                    bestWaste = waste;
                }
            }
        }
        if (bestReachIdx >= 0) {
            return bestReachIdx;
        }
        return pool.isEmpty() ? -1 : 0;
    }

    private static BigDecimal sumAmounts(List<InvoiceRecord> invoices) {
        return invoices.stream()
                .map(InvoiceRecord::totalTaxIncludedAmount)
                .reduce(BigDecimal.ZERO, BigDecimal::add);
    }

    private static void writeWorkbook(List<GroupRow> groupedRows,
                                      List<GroupRow> insufficientRows,
                                      Path output) throws Exception {
        Path parent = output.toAbsolutePath().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet1 = wb.createSheet("sheet1_凑100");
            Sheet sheet2 = wb.createSheet("sheet2_不足100");

            Styles styles = buildStyles(wb);
            writeGroupedSheet(sheet1, groupedRows, styles);
            writeGroupedSheet(sheet2, insufficientRows, styles);

            try (OutputStream out = Files.newOutputStream(output)) {
                wb.write(out);
            }
        }
    }

    private static void writeGroupedSheet(Sheet sheet, List<GroupRow> groupRows, Styles styles) {
        int rowCursor = 0;
        int blockIndex = 0;
        String lastSeller = null;
        for (GroupRow group : groupRows) {
            if (rowCursor > 0) {
                rowCursor++;
            }
            if (lastSeller != null && !lastSeller.equals(group.sellerIdNum())) {
                rowCursor++;
            }
            lastSeller = group.sellerIdNum();
            StylePair pair = styles.forIndex(blockIndex++);

            Row h = sheet.createRow(rowCursor++);
            for (int c = 0; c < GROUP_HEADERS.length; c++) {
                Cell cell = h.createCell(c);
                cell.setCellValue(GROUP_HEADERS[c]);
                cell.setCellStyle(pair.headerStyle());
            }

            int detailStart = rowCursor;
            int detailLines = 0;
            List<DetailLine> detailLinesOrdered = buildDetailLinesSortedByIssueTime(group.invoices());
            for (DetailLine line : detailLinesOrdered) {
                InvoiceRecord inv = line.invoice();
                Row r = sheet.createRow(rowCursor++);
                detailLines++;
                for (int c = 0; c < GROUP_HEADERS.length; c++) {
                    r.createCell(c).setCellStyle(pair.bodyStyle());
                }
                r.getCell(2).setCellValue(inv.taxBureauName());
                r.getCell(3).setCellValue(inv.invoiceNumber());
                r.getCell(4).setCellValue(inv.issueTime());
                r.getCell(5).setCellValue(inv.sellerIdNum());
                r.getCell(6).setCellValue(inv.sellerName());
                r.getCell(7).setCellValue(line.item());
                r.getCell(8).setCellValue(inv.totalTaxIncludedAmount().doubleValue());
            }
            int detailEnd = detailStart + Math.max(detailLines - 1, 0);
            writeMergedText(sheet, 0, detailStart, detailEnd, "G" + group.groupNo(), pair.bodyStyle());
            writeMergedText(sheet, 1, detailStart, detailEnd, group.groupTag(), pair.bodyStyle());
            writeMergedNumber(sheet, 9, detailStart, detailEnd, group.groupSum(), pair.bodyStyle());
        }
        autoWidth(sheet, GROUP_HEADERS.length);
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
        Arrays.fill(maxUnits, 10);
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
            pairs.add(new StylePair(createStyle(wb, BLOCK_HEADER_COLORS[i], true),
                    createStyle(wb, BLOCK_BODY_COLORS[i], false)));
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

    private static String cellString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                double d = cell.getNumericCellValue();
                if (d == Math.rint(d) && !Double.isInfinite(d)) {
                    yield String.valueOf((long) d);
                }
                yield String.valueOf(d);
            }
            default -> "";
        };
    }

    private static boolean isInvoiceNumber(String value) {
        if (value == null || value.isBlank()) {
            return false;
        }
        String v = value.trim();
        return v.length() >= 15 && v.chars().allMatch(Character::isDigit);
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

    private static String normItem(String item) {
        if (item == null) {
            return "";
        }
        return item.replaceAll("\\s+", "").trim();
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

    private record InvoiceRecord(String taxBureauName,
                                 String invoiceNumber,
                                 String issueTime,
                                 String sellerIdNum,
                                 String sellerName,
                                 List<String> items,
                                 BigDecimal totalTaxIncludedAmount,
                                 Path sourcePath) {
    }

    private record HalfEntry(long sum, int mask) {
    }

    private static final class SubsetHolder {
        long bestMask = 0L;
        long bestWaste = Long.MAX_VALUE;
    }

    private record SellerPackResult(String sellerId, PackResult packResult) {
    }

    private record GroupRow(int groupNo,
                            String groupTag,
                            List<InvoiceRecord> invoices,
                            BigDecimal groupSum,
                            String sellerIdNum) {
    }

    private record PackResult(List<List<InvoiceRecord>> groups, List<InvoiceRecord> leftovers) {
    }

    private record StylePair(CellStyle headerStyle, CellStyle bodyStyle) {
    }

    private record Styles(List<StylePair> pairs) {
        StylePair forIndex(int index) {
            return pairs.get(index % pairs.size());
        }
    }
}
