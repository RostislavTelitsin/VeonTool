module com.example.veontool {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;


    opens com.example.veontool to javafx.fxml;
    exports com.example.veontool;
}