package utils;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import javax.swing.filechooser.FileSystemView;

import bean.SelectFile;
import bean.WordInfo;

/**
 * 写整合后的Excel文件
 */
public class ExcelWriter {
    private static SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmm");

    //<editor-fold desc:"Word to Excel">
    public static void writeExcel(ArrayList<WordInfo> allWordInfoList) throws IOException {

        FileSystemView fsv = FileSystemView.getFileSystemView();
        File homeDirectory = fsv.getHomeDirectory();

        File excelFile = new File(homeDirectory + File.separator + "订单汇总(" + format.format(new Date()) + ").xlsx");
        if (!excelFile.exists()) {
            excelFile.createNewFile();
        }
        int rowCount = 0;
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("明细");
        //写表头
        CellStyle titleCellStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        titleCellStyle.setFont(font);
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);

        Row rowTitle = sheet.createRow(rowCount);
        for (int i = 0; i < 14; i++) {
            Cell cell = rowTitle.createCell(i);
            cell.setCellStyle(titleCellStyle);

            switch (i) {
                case 0:
                    cell.setCellValue("序号");
                    break;
                case 1:
                    cell.setCellValue("文件源");
                    break;
                case 2:
                    cell.setCellValue("订单编号");
                    break;
                case 3:
                    cell.setCellValue("供货单号");
                    break;
                case 4:
                    cell.setCellValue("供电局");
                    break;
                case 5:
                    cell.setCellValue("项目名称");
                    break;
                case 6:
                    cell.setCellValue("货物名称");
                    break;
                case 7:
                    cell.setCellValue("型号");
                    break;
                case 8:
                    cell.setCellValue("数量");
                    break;
                case 9:
                    cell.setCellValue("单位");
                    break;
                case 10:
                    cell.setCellValue("不含");
                    break;
                case 11:
                    cell.setCellValue("不含总价");
                    break;
                case 12:
                    cell.setCellValue("总金额");
                    break;
                case 13:
                    cell.setCellValue("交货时间");
                    break;
            }
        }

        rowCount = 1;

        //内容整理
        //Collections.sort(allWordInfoList, new WordInfo.OrderComparator());
        Iterator<WordInfo> iterable = allWordInfoList.iterator();
        String orderNumber = iterable.next().getOrderNumber();
        while (iterable.hasNext()) {
            WordInfo wordInfo = iterable.next();
            if (orderNumber.equals(wordInfo.getOrderNumber()))
                wordInfo.setOrderNumber("");
            else orderNumber = wordInfo.getOrderNumber();
        }

        for (int i = 0; i < allWordInfoList.size(); i++) {
            System.out.println("Ordernumber: (" + i + ")  " + allWordInfoList.get(i).getOrderNumber());
        }

        //填充内容
        CellStyle centerStyle = sheet.getWorkbook().createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        //centerStyle.setFont(font);
        //centerStyle.setWrapText(true);
        for (WordInfo wordInfo : allWordInfoList) {
            int columnCount = 0;
            Row row = sheet.createRow(rowCount++);

            Cell cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(rowCount - 1);//id
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getSrcfile());//文件名
            //cell.setCellValue(wordInfo.getOrderNumber().length() == 0. ? "" : wordInfo.getOrderNumber());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getOrderNumber());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getSupplyOrderNumber());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(formatPowerSupply(wordInfo.getPowerSupply()));//省略部分文字
            cell = row.createCell(columnCount++);
            cell.setCellValue(wordInfo.getProjectName());
            cell = row.createCell(columnCount++);
            cell.setCellValue(wordInfo.getGoodsName());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getModelName());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getGoodsCloums());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getGoodsUnit());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getNotwithPrice());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getNotwithTotalPrice());
            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getTotalPrice());
            cell = row.createCell(columnCount);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(wordInfo.getDeliveryTime());
        }
        FileOutputStream os = new FileOutputStream(excelFile);
        workbook.write(os);
        os.close();
    }

    public static String getFileNameWithoutEnd(String filePath) {
        File file = new File(filePath);
        String fileName = file.getName();
        if (fileName.contains(".")) {
            String[] sps = fileName.split("\\.");
            String end = sps[sps.length - 1];
            return fileName.replaceAll("\\." + end, "");
        } else {
            return fileName;
        }
    }

    private static String formatPowerSupply(String powersupply) {
        /*if (powersupply.contains("天津市电力公司") && powersupply.contains("供电分公司")) {
            return powersupply.replaceAll("天津市电力公司", "")
                    .replaceAll("供电分公司", "");
        } else return powersupply;*/
        try {
            return powersupply.split("天津市电力公司")[1].split("供电分公司")[0];
        } catch (Exception e) {
            return powersupply;
        }
    }

    //</editor-fold>

    //<editor-fold desc:"PDF to Excel">
    public static void writePDFExcel(ArrayList<SelectFile> mSelectFileList) throws IOException {
        FileSystemView fsv = FileSystemView.getFileSystemView();
        File homeDirectory = fsv.getHomeDirectory();

        File excelFile = new File(homeDirectory + File.separator + "PDF订单汇总(" + format.format(new Date()) + ").xlsx");
        if (!excelFile.exists()) {
            excelFile.createNewFile();
        }
        int rowCount = 0;
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("PDF");
        //写表头
        CellStyle titleCellStyle = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        titleCellStyle.setFont(font);
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);

        Row rowTitle = sheet.createRow(rowCount);
        for (int i = 0; i < 14; i++) {
            Cell cell = rowTitle.createCell(i);
            cell.setCellStyle(titleCellStyle);
            switch (i) {
                case 0:
                    cell.setCellValue("序号");
                    break;
                case 1:
                    cell.setCellValue("文件源");
                    break;
                case 2:
                    cell.setCellValue("订单编号");
                    break;
            }
        }
        rowCount = 1;

        //填充内容
        CellStyle centerStyle = sheet.getWorkbook().createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        for (SelectFile file : mSelectFileList) {
            int columnCount = 0;
            Row row = sheet.createRow(rowCount++);

            Cell cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(rowCount - 1);//序号

            cell = row.createCell(columnCount++);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(new File(file.getPath()).getName());//文件源

            cell = row.createCell(columnCount);
            cell.setCellStyle(centerStyle);
            cell.setCellValue(getOrderNumber(file.getName()));
        }
        FileOutputStream os = new FileOutputStream(excelFile);
        workbook.write(os);
        os.close();
    }

    private static String getOrderNumber(String filename) {
        String name = filename.split("\\.")[0];
        return name.substring(0, name.length() - 3);
    }
    //</editor-fold>
}
