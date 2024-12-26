package com.word;

import com.word.controller.WordToMarkdownController;
import javafx.application.Application;
import javafx.stage.Stage;

public class WordToMarkdownApp extends Application {

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Word to Markdown Converter");
        WordToMarkdownController controller = new WordToMarkdownController(primaryStage);
        controller.init();
    }

    public static void main(String[] args) {
        launch(args);
    }
}