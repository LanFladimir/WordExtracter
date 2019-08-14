package main;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;

import bean.SelectFile;
import bean.WordInfo;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.stage.FileChooser;
import utils.ExcelWriter;
import utils.WordReader;

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
    private Button bt_chooser_word;
    @FXML
    private Button bt_chooser_pdf;

    /**
     * chooce files(Office Word/PDF)
     */
    private void chooseFiles(int type) {
        FileChooser docChooser = new FileChooser();

        if (type == 0) {
            docChooser.setTitle("选择待筛选Word文档");
            docChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Word Files", "*.doc", "*.docx"),
                    new FileChooser.ExtensionFilter("All Files", "*.*"));
        } else {
            docChooser.setTitle("选择PDF文档");
            docChooser.getExtensionFilters().addAll(
                    new FileChooser.ExtensionFilter("Word Files", "*.pdf"),
                    new FileChooser.ExtensionFilter("All Files", "*.*"));
        }
        if (wordFileLastPath.length() != 0)
            docChooser.setInitialDirectory(new File(wordFileLastPath));
        List<File> list = docChooser.showOpenMultipleDialog(bt_chooser_word.getScene().getWindow());
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
        } else {
            new Thread(() -> {
                if (mSelectFileList.size() > 0) {
                    Platform.runLater(() -> bt_extracter.setText("导出中..."));
                    try {
                        ExcelWriter.writePDFExcel(mSelectFileList);
                        Platform.runLater(() -> bt_extracter.setText("成功"));
                    } catch (IOException e) {
                        e.printStackTrace();
                        Platform.runLater(() -> bt_extracter.setText("失败"));
                    }

                } else {
                    Platform.runLater(() -> bt_extracter.setText("无PDF"));
                }
                try {
                    Thread.sleep(1000 * 2);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                Platform.runLater(() -> bt_extracter.setText("无PDF"));
            }).start();
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

        bt_chooser_word.setOnAction(event -> chooseFiles(0));
        bt_chooser_pdf.setOnAction(event -> chooseFiles(1));
        bt_extracter.setOnAction(event -> doExtracter());
    }
}
