package com.word.service;

import com.word.strategy.ConverterStrategy;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Logger;

public class ConversionService {

    private static final Logger logger = Logger.getLogger(ConversionService.class.getName());

    private ConverterStrategy converterStrategy;

    public void setConverterStrategy(ConverterStrategy converterStrategy) {
        this.converterStrategy = converterStrategy;
    }

    public String convert(File inputFile, File outputDir) throws IOException {
        return converterStrategy.convert(inputFile, outputDir);
    }

    public static void saveFile(String content, String originalFileName, String extension) throws IOException {
        String outputFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.')) + extension;
        File outputFile = new File(outputFileName);
        try (FileWriter writer = new FileWriter(outputFile)) {
            writer.write(content);
            logger.info("File saved: " + outputFileName);
        }
    }
}