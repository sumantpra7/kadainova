package com.cyberdeveloper.assignmentmaker.controller;

import com.cyberdeveloper.assignmentmaker.service.DocumentService;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import jakarta.servlet.http.HttpServletRequest;

import java.io.IOException;
import java.nio.file.*;
import java.util.*;

/**
 * Spring MVC controller — maps all Flask routes 1:1.
 */
@Controller
public class AssignmentController {

    private final DocumentService documentService;

    public AssignmentController(DocumentService documentService) {
        this.documentService = documentService;
    }

    // ─── GET / ────────────────────────────────────────────────────────

    /**
     * Home page: detect placeholders and render form.
     * Equivalent to Flask @app.route("/")
     */
    @GetMapping("/")
    public String index(Model model) {
        List<String> placeholders = documentService.detectPlaceholders();
        model.addAttribute("placeholders", placeholders);
        model.addAttribute("error", placeholders.isEmpty()
                ? "No placeholders detected in DOCX template." : null);
        model.addAttribute("pdfReady", false);
        return "index";
    }

    // ─── POST /generate ───────────────────────────────────────────────

    /**
     * Generate PDF from form data.
     * Equivalent to Flask @app.route("/generate", methods=["POST"])
     */
    @PostMapping("/generate")
    public String generate(
            @RequestParam(value = "logoFile", required = false) MultipartFile logoFile,
            @RequestParam(value = "addBorder", defaultValue = "false") String addBorder,
            HttpServletRequest request,
            Model model) {

        List<String> placeholders = documentService.detectPlaceholders();

        // Collect form values (same as Python: context[ph] = request.form.get(ph, ""))
        Map<String, String> context = new LinkedHashMap<>();
        for (String ph : placeholders) {
            context.put(ph, request.getParameter(ph) != null ? request.getParameter(ph) : "");
        }

        // Handle logo upload
        String logoPath = null;
        if (logoFile != null && !logoFile.isEmpty()) {
            try {
                String originalName = logoFile.getOriginalFilename();
                String ext = ".png";
                if (originalName != null && originalName.contains(".")) {
                    ext = originalName.substring(originalName.lastIndexOf('.'));
                }
                Path savePath = Path.of(documentService.getUploadDir(), "custom_logo" + ext);
                logoFile.transferTo(savePath.toFile());
                logoPath = savePath.toString();
            } catch (IOException e) {
                System.err.println("WARNING: Could not save logo: " + e.getMessage());
            }
        }

        // Fill DOCX → Convert to PDF
        try {
            String filledDocx = documentService.fillDocx(context, logoPath);
            documentService.convertToPdf(filledDocx);

            // Add border if requested
            if ("true".equalsIgnoreCase(addBorder)) {
                documentService.addBorderToPdf(
                        documentService.getOutputPdfPath().toString());
            }

            model.addAttribute("placeholders", placeholders);
            model.addAttribute("pdfReady", true);

        } catch (Exception e) {
            model.addAttribute("placeholders", placeholders);
            model.addAttribute("error", "PDF generation failed: " + e.getMessage());
            model.addAttribute("pdfReady", false);
        }

        return "index";
    }

    // ─── GET /download ────────────────────────────────────────────────

    /**
     * Download the generated PDF.
     * Equivalent to Flask @app.route("/download")
     */
    @GetMapping("/download")
    public ResponseEntity<Resource> download() {
        Path pdfPath = documentService.getOutputPdfPath();
        if (!Files.exists(pdfPath)) {
            return ResponseEntity.notFound().build();
        }

        Resource resource = new FileSystemResource(pdfPath.toFile());
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_PDF)
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        "attachment; filename=\"assignment.pdf\"")
                .body(resource);
    }

    // ─── GET /view-pdf ────────────────────────────────────────────────

    /**
     * Serve the PDF inline for iframe preview.
     * Equivalent to Flask @app.route("/view-pdf")
     */
    @GetMapping("/view-pdf")
    public ResponseEntity<Resource> viewPdf() {
        Path pdfPath = documentService.getOutputPdfPath();
        if (!Files.exists(pdfPath)) {
            return ResponseEntity.notFound().build();
        }

        Resource resource = new FileSystemResource(pdfPath.toFile());
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_PDF)
                .body(resource);
    }
}
