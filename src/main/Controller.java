package main;

import bean.SelectFile;
import bean.WordInfo;
import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.stage.FileChooser;
import utils.ExcelWriter;
import utils.WordReader;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Array;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

public class Controller implements Initializable {
    private String wordFileLastPath = "";
    private int fileType;

    private ObservableList<SelectFile> mTableViewData;
    private ArrayList<SelectFile> mSelectFileList = new ArrayList<>();

    @FXML
    private Button bt_extracter;

    @FXML
    private TableView<SelectFile> tv_filelist;

    @FXML
    private Button bt_chooser;

    @FXML
    private ComboBox<String> cb_option;
    private ObservableList<String> cbInfoList = FXCollections.observableArrayList("Word-->Excel", "PDF-->Excel", "Excel-->Excel");
    @FXML
    private Label lb_option;
    private String optionType = "Word-->Excel";

    /**
     * chooce files(Office Word/PDF)
     */
    private void chooseFiles(int type) {
        FileChooser docChooser = new FileChooser();

        if (type == 0) {
            docChooser.setTitle("选择待解析Word文档");
            docChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Word Files", "*.doc", "*.docx"),
                    new FileChooser.ExtensionFilter("All Files", "*.*"));
        } else if (type == 1) {
            docChooser.setTitle("选择PDF文档");
            docChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Word Files", "*.pdf"),
                    new FileChooser.ExtensionFilter("All Files", "*.*"));
        } else if (type == 2) {
            //todo Excel to Excel
            docChooser.setTitle("选择待解析Excel文档");
            docChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Word Files", "*.xlsx"),
                    new FileChooser.ExtensionFilter("All Files", "*.*"));
        } else {
            System.out.println("not match method");
        }
        if (wordFileLastPath.length() != 0)
            docChooser.setInitialDirectory(new File(wordFileLastPath));
        List<File> list = docChooser.showOpenMultipleDialog(tv_filelist.getScene().getWindow());
        if (list != null) {
            fileType = type;
            mSelectFileList.clear();

            list.forEach((file) -> {
                mSelectFileList.add(new SelectFile(file.getName(), file.getAbsolutePath()));
                System.out.println(file.getName());
            });
        }
        mTableViewData.clear();
        mTableViewData.addAll(mSelectFileList);
    }

    /**
     * extracter infos and make new Excel file
     */
    private void doExtracter() {
        if (fileType == 0) {
            if (mSelectFileList.size() > 0) {
                bt_extracter.setText("解析中...");
                new Thread(() -> {
                    ArrayList<WordInfo> allWordInfoList = new ArrayList<>();
                    for (SelectFile wordFile : mSelectFileList) {
                        allWordInfoList.addAll(WordReader.readWord(wordFile.getPath()));
                    }
                    try {
                        ExcelWriter.writeExcel(allWordInfoList);
                        setExtracterCallBack(0);
                    } catch (IOException e) {
                        e.printStackTrace();
                        setExtracterCallBack(1);
                        Platform.runLater(() -> bt_extracter.setText(""));
                    }
                }).start();
            } else {
                setExtracterCallBack(2);
            }
        } else if (fileType == 1) {
            new Thread(() -> {
                if (mSelectFileList.size() > 0) {
                    Platform.runLater(() -> bt_extracter.setText("导出中..."));
                    try {
                        ExcelWriter.writePDFExcel(mSelectFileList);
                        setExtracterCallBack(0);
                    } catch (IOException e) {
                        e.printStackTrace();
                        setExtracterCallBack(1);
                    }

                } else {
                    setExtracterCallBack(2);
                }
            }).start();
        } else if (fileType == 2) {
            new Thread(() -> {
                if (mSelectFileList.size() > 0) {
                    Platform.runLater(() -> bt_extracter.setText("解析中..."));
                    try {
                        ExcelWriter.ExcelToExcel(mSelectFileList);
                        setExtracterCallBack(10);
                    } catch (IOException e) {
                        e.printStackTrace();
                        setExtracterCallBack(1);
                    }
                } else {
                    setExtracterCallBack(2);
                }
            }).start();
        } else {
            System.out.println("oooops");
        }
    }

    /**
     * change UI
     */
    private void setExtracterCallBack(int state) {
        switch (state) {
            case 0:
                Platform.runLater(() -> bt_extracter.setText("解析完成"));
                break;
            case 1:
                Platform.runLater(() -> bt_extracter.setText("解析失败"));
                break;
            case 2:
                Platform.runLater(() -> bt_extracter.setText("无文件"));
                break;
        }
        new Thread(() -> {
            try {
                Thread.sleep(1000 * 2);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            Platform.runLater(() -> bt_extracter.setText("解析"));
        }).start();
    }

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        System.out.println("init");
        mTableViewData = FXCollections.observableArrayList(mSelectFileList);
        TableColumn fileCol = new TableColumn("文件名");
        fileCol.setCellValueFactory(new PropertyValueFactory<>("name"));
        TableColumn pathCol = new TableColumn("文件路径");
        pathCol.setCellValueFactory(new PropertyValueFactory<>("path"));
        mTableViewData = FXCollections.observableArrayList(mSelectFileList);
        tv_filelist.setItems(mTableViewData);
        tv_filelist.getColumns().addAll(fileCol, pathCol);

        cb_option.setItems(cbInfoList);
        cb_option.setPromptText("选择功能类型");
        cb_option.setValue("Word-->Excel");
        lb_option.setText("WORD批量提取数据导出Excel");
        bt_chooser.setOnAction(event -> chooseFiles(0));

        cb_option.valueProperty().addListener((observable, oldValue, newValue) -> {
            System.out.println("选择了" + newValue);
            optionType = newValue;

            switch (optionType) {
                case "Word-->Excel":
                    bt_chooser.setOnAction(event -> chooseFiles(0));
                    lb_option.setText("WORD批量提取数据导出Excel");
                    break;
                case "PDF-->Excel":
                    bt_chooser.setOnAction(event -> chooseFiles(1));
                    lb_option.setText("PDF文件名导出至Excel");
                    break;
                case "Excel-->Excel":
                    bt_chooser.setOnAction(event -> chooseFiles(2));
                    lb_option.setText("Excel批量提取数据导出Excel");
                    break;
            }

            //数据清空
            mSelectFileList.clear();
            mTableViewData.clear();
        });


        bt_extracter.setOnAction(event -> doExtracter());
    }

}
