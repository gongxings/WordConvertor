package com.word.converter;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class WordToMarkdownConverter {
    private static final Logger logger = Logger.getLogger(WordToMarkdownConverter.class.getName());

    public static String convertDocToMarkdown(File docFile, File imageDir) throws IOException {
        StringBuilder markdown = new StringBuilder();
        logger.info("Starting conversion of .doc file: " + docFile.getName());
        try (FileInputStream fis = new FileInputStream(docFile)) {
            HWPFDocument doc = new HWPFDocument(fis);

            Range docRange = doc.getRange();

            // Process paragraphs and tables
            int numParagraphs = docRange.numParagraphs();
            for (int i = 0; i < numParagraphs; i++) {
                Paragraph paragraph = docRange.getParagraph(i);
                try {
                    // Skip table of contents paragraphs
                    if (isTableOfContents(paragraph.text())) {
                        logger.info("Skipping Table of Contents paragraph: " + i);
                        continue;
                    }
                    if (paragraph.isInTable()) {
                        if(isTableStart(paragraph, doc, i)){
                            Table table = docRange.getTable(paragraph);
                            markdown.append(convertTableToMarkdown(table));
                            while (i < numParagraphs && docRange.getParagraph(i).isInTable()) {
                                i++;
                            }
                            i--; // Adjust for the extra increment in the while loop
                            logger.info("Processed table at paragraph: " + i);
                        }
                    } else {
                        markdown.append(processParagraph(paragraph, doc, imageDir)).append("\n");
                        logger.info("Processed paragraph: " + i);
                    }
                } catch (Exception e) {
                    logger.log(Level.SEVERE, "Error processing paragraph: " + i, e);
                }
            }
        }
        logger.info("Completed conversion of .doc file: " + docFile.getName());
        return markdown.toString();
    }

    private static boolean isTableStart(Paragraph paragraph, HWPFDocument doc, int paragraphIndex) {
        if (!paragraph.isInTable()) {
            return false;
        }
        if (paragraphIndex == 0) {
            return true;
        }
        Paragraph previousParagraph = doc.getRange().getParagraph(paragraphIndex - 1);
        return !previousParagraph.isInTable();
    }

    private static boolean isTableEnd(Paragraph paragraph, HWPFDocument doc, int paragraphIndex) {
        if (!paragraph.isInTable()) {
            return false;
        }
        if (paragraphIndex == doc.getRange().numParagraphs() - 1) {
            return true;
        }
        Paragraph nextParagraph = doc.getRange().getParagraph(paragraphIndex + 1);
        return !nextParagraph.isInTable();
    }

    public static String convertDocxToMarkdown(File docxFile, File imageDir) throws IOException {
        StringBuilder markdown = new StringBuilder();
        logger.info("Starting conversion of .docx file: " + docxFile.getName());
        try (FileInputStream fis = new FileInputStream(docxFile)) {
            XWPFDocument docx = new XWPFDocument(fis);

            // Process paragraphs and tables
            List<IBodyElement> bodyElements = docx.getBodyElements();
            for (int i = 0; i < bodyElements.size(); i++) {
                IBodyElement element = bodyElements.get(i);
                try {
                    if (element instanceof XWPFParagraph) {
                        XWPFParagraph paragraph = (XWPFParagraph) element;
                        // Skip table of contents paragraphs
                        if (isTableOfContents(paragraph.getText())) {
                            logger.info("Skipping Table of Contents paragraph: " + i);
                            continue;
                        }
                        markdown.append(processParagraph(paragraph, docx, imageDir)).append("\n");
                        logger.info("Processed paragraph: " + i);
                    } else if (element instanceof XWPFTable) {
                        XWPFTable table = (XWPFTable) element;
                        markdown.append(convertTableToMarkdown(table));
                        logger.info("Processed table: " + i);
                    }
                } catch (Exception e) {
                    logger.log(Level.SEVERE, "Error processing element: " + i, e);
                }
            }
        }
        logger.info("Completed conversion of .docx file: " + docxFile.getName());
        return markdown.toString();
    }

    private static boolean isTableOfContents(String text) {
        // Check for common table of contents keywords and patterns
        return text.trim().matches("^[\\d\\.\\s]*目录$") ||
               text.contains("TOC \\o") ||
               text.contains("PAGEREF");
    }

    private static String processParagraph(Paragraph paragraph, HWPFDocument doc, File imageDir) throws IOException {
        StringBuilder result = new StringBuilder();
        if (paragraph.getStyleIndex() != 0) {
            // Convert headings
            int headingLevel = paragraph.getStyleIndex();
            result.append("#".repeat(Math.max(0, headingLevel))).append(" ").append(paragraph.text().trim());
        } else {
            result.append(paragraph.text().trim());
        }
        
        // Handle images within the paragraph
        int numCharacterRuns = paragraph.numCharacterRuns();
        for (int j = 0; j < numCharacterRuns; j++) {
            CharacterRun characterRun = paragraph.getCharacterRun(j);
            Picture picture = doc.getPicturesTable().extractPicture(characterRun, false);
            if (picture != null) {
                String ext = picture.suggestFileExtension();
                if (ext.isEmpty()) {
                    ext = "png"; // Default extension
                }else if(ext.equals("emf")){
                    ext = "jpg";
                }
                String fullFileName = picture.suggestFullFileName();
                String fileNameWithoutExtension = fullFileName.contains(".") ? fullFileName.substring(0, fullFileName.lastIndexOf('.')) : fullFileName;
                String imageFileName = "image-" + fileNameWithoutExtension + "." + ext;
                File imageFile = new File(imageDir, imageFileName);
                try (FileOutputStream fos = new FileOutputStream(imageFile)) {
                    picture.writeImageContent(fos);
                }
                result.append("\n![Image](").append(imageFile.getPath()).append(")");
            } else if (characterRun.text().contains("EMBED Visio.Drawing.11")) {
                // Handle special case for Visio drawings
                result.append("\n![Visio Drawing](path/to/visio/drawing.png)");
            }
        }
        return result.toString();
    }

    private static String processParagraph(XWPFParagraph paragraph, XWPFDocument docx, File imageDir) throws IOException {
        StringBuilder result = new StringBuilder();
        if (paragraph.getStyle() != null && paragraph.getStyle().startsWith("Heading")) {
            // Convert headings
            int headingLevel = Integer.parseInt(paragraph.getStyle().substring("Heading".length()));
            result.append("#".repeat(Math.max(0, headingLevel))).append(" ").append(paragraph.getText().trim());
        } else {
            result.append(paragraph.getText().trim());
        }

        // Handle images within the paragraph
        for (XWPFRun run : paragraph.getRuns()) {
            for (XWPFPicture pic : run.getEmbeddedPictures()) {
                XWPFPictureData pictureData = pic.getPictureData();
                String ext = pictureData.suggestFileExtension();
                if (ext.isEmpty()) {
                    ext = "png"; // Default extension
                }
                String imageFileName = "image" + docx.getAllPictures().indexOf(pictureData) + "." + ext;
                File imageFile = new File(imageDir, imageFileName);
                Files.write(imageFile.toPath(), pictureData.getData());
                result.append("\n![Image](").append(imageFile.getPath()).append(")");
            }
            if (run.getText(0) != null && run.getText(0).contains("EMBED Visio.Drawing.11")) {
                // Handle special case for Visio drawings
                result.append("\n![Visio Drawing](path/to/visio/drawing.png)");
            } else if (run.getText(0) != null && run.getText(0).contains(".emf")) {
                // Handle .emf images
                String imageFileName = "image" + System.currentTimeMillis() + ".png";
                File imageFile = new File(imageDir, imageFileName);
                convertEmfToPng(run.getText(0), imageFile);
                result.append("\n![Image](").append(imageFile.getPath()).append(")");
            }
        }
        return result.toString();
    }

    private static void convertEmfToPng(String emfPath, File outputFile) throws IOException {
        // Convert .emf to .png using Apache Batik
        BufferedImage bufferedImage = ImageIO.read(new File(emfPath));
        ImageIO.write(bufferedImage, "png", outputFile);
    }

    private static String convertTableToMarkdown(Table table) {
        StringBuilder markdown = new StringBuilder();
        int numRows = table.numRows();
        for (int i = 0; i < numRows; i++) {
            TableRow row = table.getRow(i);
            int numCells = row.numCells();
            for (int j = 0; j < numCells; j++) {
                markdown.append("| ").append(row.getCell(j).text().trim()).append(" ");
            }
            markdown.append("|\n");
            if (i == 0) {
                for (int j = 0; j < numCells; j++) {
                    markdown.append("|---");
                }
                markdown.append("|\n");
            }
        }
        return markdown.toString();
    }

    private static String convertTableToMarkdown(XWPFTable table) {
        StringBuilder markdown = new StringBuilder();
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                markdown.append("| ").append(cell.getText().trim()).append(" ");
            }
            markdown.append("|\n");
            if (i == 0) {
                for (int j = 0; j < cells.size(); j++) {
                    markdown.append("|---");
                }
                markdown.append("|\n");
            }
        }
        return markdown.toString();
    }
}