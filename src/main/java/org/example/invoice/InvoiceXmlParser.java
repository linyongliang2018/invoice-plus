package org.example.invoice;

import org.example.invoice.model.InvoiceData;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilderFactory;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public final class InvoiceXmlParser {

    private InvoiceXmlParser() {
    }

    public static InvoiceData parse(Path xmlPath) throws Exception {
        try (InputStream in = Files.newInputStream(xmlPath)) {
            return parse(in);
        }
    }

    public static InvoiceData parse(InputStream in) throws Exception {
        Document doc = DocumentBuilderFactory.newInstance()
                .newDocumentBuilder()
                .parse(in);
        doc.getDocumentElement().normalize();

        InvoiceData invoice = new InvoiceData();
        invoice.setEiid(textOfFirst(doc, "EIid"));
        invoice.setSellerIdNum(textOfFirst(doc, "SellerIdNum"));
        invoice.setSellerName(textOfFirst(doc, "SellerName"));
        invoice.setTotalTaxIncludedAmount(textOfFirst(doc, "TotalTax-includedAmount"));
        invoice.setIssueTime(textOfFirst(doc, "IssueTime"));
        invoice.setTaxBureauName(textOfFirst(doc, "TaxBureauName"));

        NodeList items = doc.getElementsByTagName("IssuItemInformation");
        for (int i = 0; i < items.getLength(); i++) {
            Element itemEl = (Element) items.item(i);
            InvoiceData.LineItem line = new InvoiceData.LineItem();
            line.setItemName(childText(itemEl, "ItemName"));
            line.setTotalTaxIncludedAmount(childText(itemEl, "TotaltaxIncludedAmount"));
            if (isBlank(line.getItemName()) && isBlank(line.getTotalTaxIncludedAmount())) {
                continue;
            }
            invoice.getLineItems().add(line);
        }

        if (invoice.getLineItems().isEmpty()) {
            InvoiceData.LineItem placeholder = new InvoiceData.LineItem();
            placeholder.setItemName("");
            placeholder.setTotalTaxIncludedAmount("");
            invoice.getLineItems().add(placeholder);
        }

        return invoice;
    }

    public static List<InvoiceData> parseAll(List<Path> xmlPaths) throws Exception {
        List<InvoiceData> result = new ArrayList<>();
        for (Path path : xmlPaths) {
            result.add(parse(path));
        }
        return result;
    }

    private static String textOfFirst(Document doc, String tagName) {
        NodeList nodes = doc.getElementsByTagName(tagName);
        if (nodes.getLength() == 0) {
            return "";
        }
        return nodes.item(0).getTextContent().trim();
    }

    private static String childText(Element parent, String tagName) {
        NodeList nodes = parent.getElementsByTagName(tagName);
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            if (node.getParentNode() == parent) {
                return node.getTextContent().trim();
            }
        }
        if (nodes.getLength() > 0) {
            return nodes.item(0).getTextContent().trim();
        }
        return "";
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }
}
