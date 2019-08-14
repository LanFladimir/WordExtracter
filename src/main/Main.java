package main;

import java.io.IOException;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class Main extends Application {

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        try {
            Parent root = FXMLLoader.load(getClass().getResource("main.fxml"));
            primaryStage.setScene(new Scene(root));
            //primaryStage.setResizable(false);
            primaryStage.setTitle("Word 数据提取工具");
            primaryStage.show();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
