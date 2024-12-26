package com.word.controller;

import com.word.service.ConversionService;
import com.word.strategy.MarkdownToWordConverterStrategy;
import com.word.strategy.WordToMarkdownConverterStrategy;
import com.word.converter.WordToMarkdownConverter;
import javafx.stage.Stage;
import javafx.stage.FileChooser;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.concurrent.Task;

import java.io.File;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.LogRecord;
import java.util.logging.Logger;

public class WordToMarkdownController {

    private static final Logger logger = Logger.getLogger(WordToMarkdownController.class.getName());
    private Stage primaryStage;
    private TextArea logArea;
    private ConversionService conversionService;

    public WordToMarkdownController(Stage primaryStage) {
        this.primaryStage = primaryStage;
        this.conversionService = new ConversionService();
    }

    public void init() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Word Files", "*.doc", "*.docx"),
                new FileChooser.ExtensionFilter("Markdown Files", "*.md")
        );

        Label label = new Label("Select a file to convert:");
        Button selectButton = new Button("Select File");
        ProgressBar progressBar = new ProgressBar(0);
        TextArea outputArea = new TextArea();
        outputArea.setEditable(false);
        logArea = new TextArea();
        logArea.setEditable(false);

        // Add a custom log handler to display logs in the logArea
        Handler logHandler = new Handler() {
            @Override
            public void publish(LogRecord record) {
                logArea.appendText(record.getLevel() + ": " + record.getMessage() + "\n");
            }

            @Override
            public void flush() {
            }

            @Override
            public void close() throws SecurityException {
            }
        };
        Logger.getLogger(WordToMarkdownConverter.class.getName()).addHandler(logHandler);
        Logger.getLogger(WordToMarkdownController.class.getName()).addHandler(logHandler);

        selectButton.setOnAction(e -> {
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                File outputDir = new File("images");
                if (!outputDir.exists()) {
                    outputDir.mkdirs();
                }

                Task<String> task = new Task<>() {
                    @Override
                    protected String call() throws Exception {
                        String result = "";
                        try {
                            if (selectedFile.getName().endsWith(".doc") || selectedFile.getName().endsWith(".docx")) {
                                conversionService.setConverterStrategy(new WordToMarkdownConverterStrategy());
                                result = conversionService.convert(selectedFile, outputDir);
                                ConversionService.saveFile(result, selectedFile.getName(), ".md");
                            } else if (selectedFile.getName().endsWith(".md")) {
                                conversionService.setConverterStrategy(new MarkdownToWordConverterStrategy());
                                result = conversionService.convert(selectedFile, outputDir);
                                ConversionService.saveFile(result, selectedFile.getName(), ".docx");
                            }
                        } catch (Exception ex) {
                            logger.log(Level.SEVERE, "Error during conversion", ex);
                        }
                        return result;
                    }
                };

                // Bind progress bar to task progress
                progressBar.progressProperty().bind(task.progressProperty());

                task.setOnSucceeded(event -> {
                    outputArea.setText(task.getValue());
                    progressBar.progressProperty().unbind();
                    progressBar.setProgress(1.0); // Manually set to 100% on success
                });

                task.setOnFailed(event -> {
                    progressBar.progressProperty().unbind();
                    progressBar.setProgress(0); // Reset progress bar on failure
                });

                new Thread(task).start();
            }
        });

        HBox topRow = new HBox(10, selectButton, progressBar);
        VBox root = new VBox(10, label, topRow, outputArea, new Label("Log Output:"), logArea);
        Scene scene = new Scene(root, 600, 600);
        primaryStage.setScene(scene);
        primaryStage.show();
    }
}