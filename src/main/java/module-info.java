module com.word.worddemo {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.scratchpad;
    requires org.apache.poi.ooxml;
    requires org.commonmark;
    requires java.logging;
    requires java.desktop;
    exports com.word.converter;
    opens com.word.converter to javafx.fxml;
    exports com.word.service;
    opens com.word.service to javafx.fxml;
    exports com.word;
    opens com.word to javafx.fxml;
}