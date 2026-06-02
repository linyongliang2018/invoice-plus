package org.example.invoice;

import jakarta.mail.BodyPart;
import jakarta.mail.Flags;
import jakarta.mail.Folder;
import jakarta.mail.Message;
import jakarta.mail.Multipart;
import jakarta.mail.Part;
import jakarta.mail.Session;
import jakarta.mail.Store;
import jakarta.mail.internet.MimeUtility;
import jakarta.mail.FetchProfile;
import jakarta.mail.search.AndTerm;
import jakarta.mail.search.ComparisonTerm;
import jakarta.mail.search.ReceivedDateTerm;
import jakarta.mail.search.SearchTerm;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.net.URLDecoder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 批量下载 QQ 邮箱中的京东电子发票 XML。
 *
 * <p>运行方式：直接运行 main 方法（配置已写死）。</p>
 */
public class QQJdInvoiceXmlDownloader {

    private static final String IMAP_HOST = "imap.qq.com";
    private static final String IMAP_PORT = "993";
    private static final String EMAIL = "406967240@qq.com";
    private static final String AUTH_CODE = "gpnoocrfyntzbhjh";

    private static final String SUBJECT_KEYWORD = "电子发票已开具";
    private static final String SUBJECT_PREFIX_KEYWORD = "您的京东订单";
    private static final String ANCHOR_TEXT_KEYWORD = "发票XML文件下载";
    private static final Pattern URL_PATTERN = Pattern.compile("https?://[^\\s\"'<>]+");

    /** 仅扫描最近 N 小时邮件，6 = 最近 6 小时 */
    private static final int RECENT_HOURS = 6;
    /** true 时仅测试 IMAP 连通并读取全部邮件，不做下载 */
    private static final boolean READ_ALL_MAILS_ONLY = false;
    /** READ_ALL_MAILS_ONLY 模式下最多打印前 N 封主题 */
    private static final int PREVIEW_SUBJECT_COUNT = 500;
    private static final int RETRY_TIMES = 3;

    private static final Path DOWNLOAD_DIR =
            Path.of("C:\\Users\\40696\\Desktop\\京东发票-京东发票明细");

    public static void main(String[] args) throws Exception {
        Files.createDirectories(DOWNLOAD_DIR);

        Properties props = new Properties();
        props.put("mail.store.protocol", "imaps");
        props.put("mail.imaps.host", IMAP_HOST);
        props.put("mail.imaps.port", IMAP_PORT);
        props.put("mail.imaps.ssl.enable", "true");
        props.put("mail.imaps.connectiontimeout", "10000");
        props.put("mail.imaps.timeout", "15000");

        Session session = Session.getInstance(props);
        session.setDebug(false);

        int skipped = 0;
        int success = 0;
        int failed = 0;
        int noLink = 0;

        try (Store store = session.getStore("imaps")) {
            store.connect(IMAP_HOST, EMAIL, AUTH_CODE);
            Folder inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_ONLY);

            if (READ_ALL_MAILS_ONLY) {
                printLatestSubjectsQuickly(inbox);
                inbox.close(false);
                return;
            }

            Message[] messages = searchMessages(inbox);
            System.out.printf("命中邮件 %d 封，开始解析下载...%n", messages.length);

            List<Message> sorted = new ArrayList<>(List.of(messages));
            sorted.sort(Comparator.comparing(QQJdInvoiceXmlDownloader::safeReceivedDate));

            for (Message message : sorted) {
                String subject = safeSubject(message);
                if (!isJdInvoiceSubject(subject)) {
                    continue;
                }
                String html = extractPreferredHtml(message);
                if (html == null || html.isBlank() || !html.contains(ANCHOR_TEXT_KEYWORD)) {
                    continue;
                }

                String downloadUrl = findXmlDownloadUrl(html);
                if (downloadUrl == null || downloadUrl.isBlank()) {
                    noLink++;
                    System.out.printf("[无链接] 主题=%s%n", subject);
                    continue;
                }

                DownloadResult result = downloadWithRetry(downloadUrl, subject);
                if (result.status() == DownloadStatus.SKIPPED) {
                    skipped++;
                } else if (result.status() == DownloadStatus.SUCCESS) {
                    success++;
                } else {
                    failed++;
                }
            }

            inbox.close(false);
        }

        System.out.println("========== 下载完成 ==========");
        System.out.printf("成功: %d, 跳过(已存在): %d, 失败: %d, 无链接: %d%n",
                success, skipped, failed, noLink);
        System.out.printf("下载目录: %s%n", DOWNLOAD_DIR.toAbsolutePath());
    }

    private static void printLatestSubjectsQuickly(Folder inbox) throws Exception {
        int total = inbox.getMessageCount();
        System.out.printf("连接成功，INBOX 总邮件数: %d%n", total);
        if (total <= 0) {
            return;
        }

        int start = Math.max(1, total - PREVIEW_SUBJECT_COUNT + 1);
        Message[] latest = inbox.getMessages(start, total);

        FetchProfile fetchProfile = new FetchProfile();
        fetchProfile.add(FetchProfile.Item.ENVELOPE);
        inbox.fetch(latest, fetchProfile);

        int index = 1;
        Date threshold = new Date(System.currentTimeMillis() - TimeUnit.HOURS.toMillis(RECENT_HOURS));
        int matched = 0;
        for (int i = latest.length - 1; i >= 0; i--) {
            Message msg = latest[i];
            Date mailDate = safeReceivedDate(msg);
            if (mailDate.before(threshold)) {
                continue;
            }
            System.out.printf("[%d] %s | %s%n",
                    index++,
                    mailDate,
                    safeSubject(msg));
            matched++;
        }
        System.out.printf("预览窗口：最近 %d 小时，命中 %d 封（在最新 %d 封中筛选）%n",
                RECENT_HOURS, matched, latest.length);
    }

    private static Message[] searchMessages(Folder inbox) throws Exception {
        long now = System.currentTimeMillis();
        Date threshold = new Date(now - TimeUnit.HOURS.toMillis(RECENT_HOURS));

        SearchTerm dateTerm = new ReceivedDateTerm(ComparisonTerm.GE, threshold);
        SearchTerm notDeleted = new jakarta.mail.search.FlagTerm(new Flags(Flags.Flag.DELETED), false);
        SearchTerm combined = new AndTerm(dateTerm, notDeleted);
        return inbox.search(combined);
    }

    private static String extractPreferredHtml(Part part) throws Exception {
        if (part.isMimeType("text/html")) {
            Object content = part.getContent();
            return content == null ? "" : content.toString();
        }
        if (part.isMimeType("text/plain")) {
            Object content = part.getContent();
            return content == null ? "" : content.toString();
        }
        if (part.isMimeType("multipart/*")) {
            Multipart multipart = (Multipart) part.getContent();
            String plainText = null;
            for (int i = 0; i < multipart.getCount(); i++) {
                BodyPart bodyPart = multipart.getBodyPart(i);
                String contentType = bodyPart.getContentType();
                if (bodyPart.isMimeType("text/html")) {
                    String html = extractPreferredHtml(bodyPart);
                    if (html != null && !html.isBlank()) {
                        return html;
                    }
                } else if (bodyPart.isMimeType("text/plain")) {
                    if (plainText == null) {
                        plainText = extractPreferredHtml(bodyPart);
                    }
                } else if (contentType != null && contentType.toLowerCase().contains("multipart/")) {
                    String nested = extractPreferredHtml(bodyPart);
                    if (nested != null && !nested.isBlank()) {
                        return nested;
                    }
                } else if (bodyPart.isMimeType("message/rfc822")) {
                    Object nestedContent = bodyPart.getContent();
                    if (nestedContent instanceof Part nestedPart) {
                        String nested = extractPreferredHtml(nestedPart);
                        if (nested != null && !nested.isBlank()) {
                            return nested;
                        }
                    }
                }
            }
            return plainText == null ? "" : plainText;
        }
        return "";
    }

    private static String findXmlDownloadUrl(String html) {
        Document doc = Jsoup.parse(html);
        Elements anchors = doc.select("a[href]");
        for (Element anchor : anchors) {
            String text = normalize(anchor.text());
            String href = anchor.attr("abs:href").isBlank() ? anchor.attr("href") : anchor.attr("abs:href");
            String hrefLower = href == null ? "" : href.toLowerCase();
            if ((text != null && text.contains(normalize(ANCHOR_TEXT_KEYWORD)))
                    || hrefLower.contains("digital-invoice")
                    || hrefLower.contains(".xml")) {
                if (href != null && !href.isBlank()) {
                    return href;
                }
            }
        }

        String plain = normalize(doc.text());
        if (plain.contains(normalize(ANCHOR_TEXT_KEYWORD))) {
            Matcher matcher = URL_PATTERN.matcher(html);
            while (matcher.find()) {
                String candidate = matcher.group();
                String lower = candidate.toLowerCase();
                if (lower.contains("digital-invoice") || lower.contains(".xml")) {
                    return candidate;
                }
            }
        }
        return null;
    }

    private static boolean isJdInvoiceSubject(String subject) {
        String normalized = normalize(subject);
        return normalized.contains(normalize(SUBJECT_PREFIX_KEYWORD))
                && normalized.contains(normalize(SUBJECT_KEYWORD));
    }

    private static String normalize(String text) {
        if (text == null) {
            return "";
        }
        return text.replace('\u00A0', ' ')
                .replace("【", "")
                .replace("】", "")
                .replace("（", "")
                .replace("）", "")
                .replaceAll("\\s+", "")
                .trim();
    }

    private static String findXmlDownloadUrlByOriginalLogic(String html) {
        Document doc = Jsoup.parse(html);
        Elements anchors = doc.select("a[href]");
        for (Element anchor : anchors) {
            String text = anchor.text();
            if (text != null && text.contains(ANCHOR_TEXT_KEYWORD)) {
                return anchor.attr("abs:href").isBlank() ? anchor.attr("href") : anchor.attr("abs:href");
            }
        }
        return null;
    }

    private static DownloadResult downloadWithRetry(String url, String subject) {
        Exception lastError = null;
        for (int i = 1; i <= RETRY_TIMES; i++) {
            try {
                DownloadResult result = doDownload(url, subject);
                if (result.status() == DownloadStatus.FAILED) {
                    // HTTP 非 2xx，按失败重试
                    continue;
                }
                return result;
            } catch (Exception ex) {
                lastError = ex;
                System.out.printf("[重试 %d/%d] 下载异常: %s, URL=%s%n", i, RETRY_TIMES, ex.getMessage(), url);
            }
        }
        if (lastError != null) {
            System.out.printf("[下载失败] URL=%s, 错误=%s%n", url, lastError.getMessage());
        }
        return new DownloadResult(DownloadStatus.FAILED, null);
    }

    private static DownloadResult doDownload(String url, String subject) throws Exception {
        HttpURLConnection connection = null;
        try {
            URL requestUrl = URI.create(url).toURL();
            connection = (HttpURLConnection) requestUrl.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(10000);
            connection.setReadTimeout(20000);
            connection.setInstanceFollowRedirects(true);

            int code = connection.getResponseCode();
            if (code < 200 || code >= 300) {
                System.out.printf("[下载失败] HTTP=%d, URL=%s, 主题=%s%n", code, url, subject);
                return new DownloadResult(DownloadStatus.FAILED, null);
            }

            String filename = resolveFilename(connection, requestUrl);
            Path target = DOWNLOAD_DIR.resolve(filename);
            if (Files.exists(target)) {
                System.out.printf("[已跳过] 文件已存在: %s%n", target.getFileName());
                return new DownloadResult(DownloadStatus.SKIPPED, target);
            }

            Path tempFile = DOWNLOAD_DIR.resolve(filename + ".tmp");
            try (InputStream in = connection.getInputStream();
                 OutputStream out = Files.newOutputStream(tempFile,
                         StandardOpenOption.CREATE,
                         StandardOpenOption.TRUNCATE_EXISTING,
                         StandardOpenOption.WRITE)) {
                in.transferTo(out);
            }
            Files.move(tempFile, target, StandardCopyOption.REPLACE_EXISTING);
            System.out.printf("[下载成功] %s%n", target.getFileName());
            return new DownloadResult(DownloadStatus.SUCCESS, target);
        } finally {
            if (connection != null) {
                connection.disconnect();
            }
        }
    }

    private static String resolveFilename(HttpURLConnection connection, URL requestUrl) throws Exception {
        String cd = connection.getHeaderField("Content-Disposition");
        if (cd != null && !cd.isBlank()) {
            String decoded = MimeUtility.decodeText(cd);
            String marker = "filename=";
            int idx = decoded.toLowerCase().indexOf(marker);
            if (idx >= 0) {
                String raw = decoded.substring(idx + marker.length()).trim();
                raw = raw.replace("\"", "").replace("'", "");
                if (!raw.isBlank()) {
                    String name = URLDecoder.decode(raw, StandardCharsets.UTF_8);
                    return ensureXmlExt(sanitizeFilename(name));
                }
            }
        }

        String path = requestUrl.getPath();
        String fallback = path == null ? "" : path.substring(path.lastIndexOf('/') + 1);
        if (fallback.isBlank()) {
            fallback = "invoice_" + System.currentTimeMillis() + ".xml";
        }
        fallback = URLDecoder.decode(fallback, StandardCharsets.UTF_8);
        return ensureXmlExt(sanitizeFilename(fallback));
    }

    private static String ensureXmlExt(String filename) {
        String name = filename == null ? "" : filename.trim();
        if (name.isBlank()) {
            return "invoice_" + System.currentTimeMillis() + ".xml";
        }
        if (!name.toLowerCase().endsWith(".xml")) {
            return name + ".xml";
        }
        return name;
    }

    private static String sanitizeFilename(String filename) {
        return filename.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    private static Date safeReceivedDate(Message message) {
        try {
            Date d = message.getReceivedDate();
            return d == null ? new Date(0) : d;
        } catch (Exception ignored) {
            return new Date(0);
        }
    }

    private static String safeSubject(Message message) {
        try {
            String subject = message.getSubject();
            return subject == null ? "" : subject;
        } catch (Exception ignored) {
            return "";
        }
    }

    private enum DownloadStatus {
        SUCCESS,
        SKIPPED,
        FAILED
    }

    private record DownloadResult(DownloadStatus status, Path file) {
    }
}
