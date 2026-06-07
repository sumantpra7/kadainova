package com.cyberdeveloper.assignmentmaker.service;

import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import com.itextpdf.kernel.pdf.*;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.geom.Rectangle;

import jakarta.annotation.PostConstruct;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.regex.*;
import java.util.zip.*;

/**
 * Handles all DOCX template processing and PDF operations.
 * This is the Java equivalent of the Python app.py logic.
 */
@Service
public class DocumentService {

    @Value("${app.template-path}")
    private String templatePath;

    @Value("${app.output-dir}")
    private String outputDir;

    @Value("${app.upload-dir}")
    private String uploadDir;

    @Value("${app.libreoffice-path}")
    private String libreOfficePath;

    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\{\\{(.*?)}}");

    @PostConstruct
    public void init() throws IOException {
        Files.createDirectories(Path.of(outputDir));
        Files.createDirectories(Path.of(uploadDir));
    }

    // ─── Placeholder Detection ────────────────────────────────────────

    /**
     * Read the DOCX template and extract all {{placeholder}} names.
     * Equivalent to Python's detect_placeholders().
     */
    public List<String> detectPlaceholders() {
        Set<String> found = new LinkedHashSet<>();
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            // Check paragraphs in the document body
            for (XWPFParagraph para : doc.getParagraphs()) {
                extractPlaceholders(para.getText(), found);
            }

            // Also check table cells (templates often use tables)
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph para : cell.getParagraphs()) {
                            extractPlaceholders(para.getText(), found);
                        }
                    }
                }
            }

            // Check headers and footers
            for (XWPFHeader header : doc.getHeaderList()) {
                for (XWPFParagraph para : header.getParagraphs()) {
                    extractPlaceholders(para.getText(), found);
                }
            }
            for (XWPFFooter footer : doc.getFooterList()) {
                for (XWPFParagraph para : footer.getParagraphs()) {
                    extractPlaceholders(para.getText(), found);
                }
            }

        } catch (Exception e) {
            System.err.println("WARNING: Could not read template: " + e.getMessage());
        }

        List<String> sorted = new ArrayList<>(found);
        Collections.sort(sorted);
        return sorted;
    }

    private void extractPlaceholders(String text, Set<String> found) {
        if (text == null) return;
        Matcher m = PLACEHOLDER_PATTERN.matcher(text);
        while (m.find()) {
            String name = m.group(1).trim();
            if (!name.isEmpty()) {
                found.add(name);
            }
        }
    }

    // ─── Fill DOCX Template ───────────────────────────────────────────

    /**
     * Fill the DOCX template with context values and optionally replace the logo.
     * Equivalent to Python's fill_docx().
     */
    public String fillDocx(Map<String, String> context, String logoPath) throws IOException {
        Path filledDocx = Path.of(outputDir, "output_filled.docx");

        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            // Replace placeholders in body paragraphs
            for (XWPFParagraph para : doc.getParagraphs()) {
                replacePlaceholdersInParagraph(para, context);
            }

            // Replace in tables
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph para : cell.getParagraphs()) {
                            replacePlaceholdersInParagraph(para, context);
                        }
                    }
                }
            }

            // Replace in headers and footers
            for (XWPFHeader header : doc.getHeaderList()) {
                for (XWPFParagraph para : header.getParagraphs()) {
                    replacePlaceholdersInParagraph(para, context);
                }
            }
            for (XWPFFooter footer : doc.getFooterList()) {
                for (XWPFParagraph para : footer.getParagraphs()) {
                    replacePlaceholdersInParagraph(para, context);
                }
            }

            // Save the filled document
            try (FileOutputStream fos = new FileOutputStream(filledDocx.toFile())) {
                doc.write(fos);
            }
        }

        // Replace logo inside the DOCX if a custom logo was provided
        if (logoPath != null && Files.isRegularFile(Path.of(logoPath))) {
            replaceLogoInDocx(filledDocx.toString(), logoPath);
        }

        return filledDocx.toString();
    }

    /**
     * Replace {{placeholder}} text in a paragraph while preserving formatting.
     * Handles the case where Word splits placeholders across multiple runs.
     */
    private void replacePlaceholdersInParagraph(XWPFParagraph para, Map<String, String> context) {
        String fullText = para.getText();
        if (fullText == null || !fullText.contains("{{")) return;

        // Build the full text from all runs to handle split placeholders
        List<XWPFRun> runs = para.getRuns();
        if (runs == null || runs.isEmpty()) return;

        // Concatenate all run texts
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : runs) {
            String t = run.getText(0);
            if (t != null) sb.append(t);
        }

        String combined = sb.toString();
        if (!combined.contains("{{")) return;

        // Apply all replacements
        String replaced = combined;
        for (Map.Entry<String, String> entry : context.entrySet()) {
            replaced = replaced.replace("{{" + entry.getKey() + "}}", entry.getValue());
            // Handle with spaces: {{ key }}
            replaced = replaced.replace("{{ " + entry.getKey() + " }}", entry.getValue());
        }

        if (replaced.equals(combined)) return; // Nothing changed

        // Clear all existing runs and set the replaced text on the first run
        // preserving the formatting of the first run
        for (int i = runs.size() - 1; i > 0; i--) {
            para.removeRun(i);
        }
        if (!runs.isEmpty()) {
            runs.get(0).setText(replaced, 0);
        }
    }

    /**
     * Replace the logo image inside the DOCX zip structure.
     * Same zip-manipulation approach as the Python version.
     */
    private void replaceLogoInDocx(String docxPath, String logoPath) {
        try {
            Path tempDocx = Path.of(docxPath + ".tmp");
            byte[] logoBytes = Files.readAllBytes(Path.of(logoPath));

            try (ZipInputStream zin = new ZipInputStream(new FileInputStream(docxPath));
                 ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(tempDocx.toFile()))) {

                ZipEntry entry;
                while ((entry = zin.getNextEntry()) != null) {
                    zout.putNextEntry(new ZipEntry(entry.getName()));

                    if ("word/media/image1.png".equals(entry.getName())) {
                        // Replace with custom logo
                        zout.write(logoBytes);
                    } else {
                        // Copy original content
                        zin.transferTo(zout);
                    }
                    zout.closeEntry();
                }
            }

            // Replace original with temp
            Files.move(tempDocx, Path.of(docxPath), StandardCopyOption.REPLACE_EXISTING);

        } catch (Exception e) {
            System.err.println("WARNING: Could not replace logo: " + e.getMessage());
        }
    }

    // ─── PDF Conversion ───────────────────────────────────────────────

    /**
     * Convert DOCX to PDF using LibreOffice headless.
     * Equivalent to Python's convert_to_pdf().
     */
    public String convertToPdf(String docxPath) throws IOException, InterruptedException {
        Path absDocx = Path.of(docxPath).toAbsolutePath();
        Path absOutDir = Path.of(outputDir).toAbsolutePath();

        ProcessBuilder pb = new ProcessBuilder(
                libreOfficePath,
                "--headless",
                "--convert-to", "pdf",
                absDocx.toString(),
                "--outdir", absOutDir.toString()
        );
        pb.redirectErrorStream(true);

        Process process = pb.start();
        // Read output for debugging
        String output;
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
            output = String.join("\n", reader.lines().toList());
        }

        int exitCode = process.waitFor();
        if (exitCode != 0) {
            throw new IOException("LibreOffice conversion failed (exit " + exitCode + "): " + output);
        }

        return Path.of(outputDir, "output_filled.pdf").toString();
    }

    // ─── PDF Border ───────────────────────────────────────────────────

    /**
     * Draw a black rectangle border on every page of the PDF.
     * Equivalent to Python's add_border_to_pdf().
     * Uses iText 8 API.
     */
    public void addBorderToPdf(String pdfPath) {
        float margin = 28f;   // ~1cm from each edge
        float lineWidth = 1.5f;

        try {
            // Read → write to temp → replace
            Path tempPdf = Path.of(pdfPath + ".bordered.tmp");

            try (PdfReader reader = new PdfReader(pdfPath);
                 PdfWriter writer = new PdfWriter(tempPdf.toString());
                 PdfDocument pdfDoc = new PdfDocument(reader, writer)) {

                for (int i = 1; i <= pdfDoc.getNumberOfPages(); i++) {
                    PdfPage page = pdfDoc.getPage(i);
                    Rectangle pageSize = page.getPageSize();

                    PdfCanvas canvas = new PdfCanvas(page);
                    canvas.setStrokeColor(com.itextpdf.kernel.colors.ColorConstants.BLACK);
                    canvas.setLineWidth(lineWidth);
                    canvas.rectangle(
                            margin,
                            margin,
                            pageSize.getWidth() - 2 * margin,
                            pageSize.getHeight() - 2 * margin
                    );
                    canvas.stroke();
                }
            }

            // Replace original with bordered version
            Files.move(tempPdf, Path.of(pdfPath), StandardCopyOption.REPLACE_EXISTING);

        } catch (Exception e) {
            System.err.println("WARNING: Could not add border: " + e.getMessage());
        }
    }

    // ─── Utility ──────────────────────────────────────────────────────

    public String getOutputDir() {
        return outputDir;
    }

    public String getUploadDir() {
        return uploadDir;
    }

    public Path getOutputPdfPath() {
        return Path.of(outputDir, "output_filled.pdf");
    }
}
