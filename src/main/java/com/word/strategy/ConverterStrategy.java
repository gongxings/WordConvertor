package com.word.strategy;

import java.io.File;
import java.io.IOException;

public interface ConverterStrategy {
    String convert(File inputFile, File outputDir) throws IOException;
}