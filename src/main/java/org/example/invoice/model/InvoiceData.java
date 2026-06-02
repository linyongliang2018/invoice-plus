package org.example.invoice.model;

import java.util.ArrayList;
import java.util.List;

public class InvoiceData {

    private String taxBureauName;
    private String eiid;
    private String sellerIdNum;
    private String sellerName;
    private String totalTaxIncludedAmount;
    private String issueTime;
    private final List<LineItem> lineItems = new ArrayList<>();

    public String getTaxBureauName() {
        return taxBureauName;
    }

    public void setTaxBureauName(String taxBureauName) {
        this.taxBureauName = taxBureauName;
    }

    public String getEiid() {
        return eiid;
    }

    public void setEiid(String eiid) {
        this.eiid = eiid;
    }

    public String getSellerIdNum() {
        return sellerIdNum;
    }

    public void setSellerIdNum(String sellerIdNum) {
        this.sellerIdNum = sellerIdNum;
    }

    public String getSellerName() {
        return sellerName;
    }

    public void setSellerName(String sellerName) {
        this.sellerName = sellerName;
    }

    public String getTotalTaxIncludedAmount() {
        return totalTaxIncludedAmount;
    }

    public void setTotalTaxIncludedAmount(String totalTaxIncludedAmount) {
        this.totalTaxIncludedAmount = totalTaxIncludedAmount;
    }

    public String getIssueTime() {
        return issueTime;
    }

    public void setIssueTime(String issueTime) {
        this.issueTime = issueTime;
    }

    public List<LineItem> getLineItems() {
        return lineItems;
    }

    public static class LineItem {
        private String itemName;
        private String totalTaxIncludedAmount;

        public String getItemName() {
            return itemName;
        }

        public void setItemName(String itemName) {
            this.itemName = itemName;
        }

        public String getTotalTaxIncludedAmount() {
            return totalTaxIncludedAmount;
        }

        public void setTotalTaxIncludedAmount(String totalTaxIncludedAmount) {
            this.totalTaxIncludedAmount = totalTaxIncludedAmount;
        }
    }
}
