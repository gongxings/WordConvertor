package com.word.service;

import com.word.converter.WordToMarkdownConverter;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Logger;

public class WordToMarkdownService {

    private static final Logger logger = Logger.getLogger(WordToMarkdownService.class.getName());

    public static String convertToMarkdown(File file, File imageDir) throws IOException {
        if (file.getName().endsWith(".doc")) {
            return WordToMarkdownConverter.convertDocToMarkdown(file, imageDir);
        } else if (file.getName().endsWith(".docx")) {
            return WordToMarkdownConverter.convertDocxToMarkdown(file, imageDir);
        }
        throw new IllegalArgumentException("Unsupported file type: " + file.getName());
    }

    public static void saveMarkdownFile(String markdown, String originalFileName) throws IOException {
        String markdownFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.')) + ".md";
        File markdownFile = new File(markdownFileName);
        try (FileWriter writer = new FileWriter(markdownFile)) {
            writer.write(markdown);
            logger.info("Markdown file saved: " + markdownFileName);
        }
    }
}