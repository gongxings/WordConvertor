package com.word.strategy;

import com.word.converter.MarkdownToWordConverter;

import java.io.File;
import java.io.IOException;

public class MarkdownToWordConverterStrategy implements ConverterStrategy {

    @Override
    public String convert(File inputFile, File outputDir) throws IOException {
        return MarkdownToWordConverter.convertMarkdownToWord(inputFile, outputDir);
    }
}