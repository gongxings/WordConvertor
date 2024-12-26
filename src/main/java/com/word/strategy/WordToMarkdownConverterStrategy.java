package com.word.strategy;

import com.word.converter.WordToMarkdownConverter;

import java.io.File;
import java.io.IOException;

public class WordToMarkdownConverterStrategy implements ConverterStrategy {

    @Override
    public String convert(File inputFile, File outputDir) throws IOException {
        if (inputFile.getName().endsWith(".doc")) {
            return WordToMarkdownConverter.convertDocToMarkdown(inputFile, outputDir);
        } else if (inputFile.getName().endsWith(".docx")) {
            return WordToMarkdownConverter.convertDocxToMarkdown(inputFile, outputDir);
        }
        throw new IllegalArgumentException("Unsupported file type: " + inputFile.getName());
    }
}