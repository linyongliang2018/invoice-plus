# Changelog

All notable changes to this project are documented in this file.

## [2026-06-02]

### Added
- Added `InvoiceImportApp` to batch parse invoice XML files and export to Excel.
- Added invoice model and parsing/export pipeline (`InvoiceData`, `InvoiceXmlParser`, `InvoiceExcelExporter`).
- Added `QQJdInvoiceXmlDownloader` for downloading and processing JD/QQ invoice sources.
- Added `InvoiceFolderWorkbookApp` for folder-based workbook workflows.
- Added `PdfToImageBatchApp` for batch PDF-to-image conversion.

### Changed
- Upgraded `pom.xml` to a complete Maven Java 17 project setup with build plugins.
- Added dependencies: Apache POI, PDFBox, Jakarta Mail, and Jsoup.
- Updated PDF conversion workflow:
  - Group outputs by configured person keywords (extensible).
  - Output only one level of group folder (no per-PDF subfolder).
  - Merge multi-page PDFs into one vertically stitched long image (`<pdfName>.png`).

### Notes
- PDF grouping is keyword-based full-text matching (not buyer-field extraction).
- Unmatched PDFs are exported to `未匹配人员`.
