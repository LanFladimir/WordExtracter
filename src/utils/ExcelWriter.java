package utils;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.swing.filechooser.FileSystemView;

import bean.SelectFile;
import bean.WordInfo;

import static java.sql.Types.NUMERIC;

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

    //<editor-fold desc:"Excel to Excel">

    public static void ExcelToExcel(ArrayList<SelectFile> mSelectFileList) throws IOException {
        //mk new excel file
        File outputExcelFile = new File(FileSystemView.getFileSystemView().getHomeDirectory() + File.separator + "台账整理(" + format.format(new Date()) + ").xlsx");
        Workbook wirteBook = new XSSFWorkbook();
        Sheet outSheet = wirteBook.createSheet("台账");
        Row header = outSheet.createRow(0);

        CellStyle headerStyle = wirteBook.createCellStyle(); // 表头单元格样式
        XSSFFont font = ((XSSFWorkbook) wirteBook).createFont(); // 字体样式
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("台账信息");//8列
        headerCell.setCellStyle(headerStyle);
        CellRangeAddress cra = new CellRangeAddress(0, 0, 0, 8); // 起始行, 终止行, 起始列, 终止列
        outSheet.addMergedRegion(cra);

        Row firstLine = outSheet.createRow(1);
        Cell titleCell = firstLine.createCell(0);
        titleCell.setCellValue("序号");
        titleCell = firstLine.createCell(1);
        titleCell.setCellValue("采购订单号");
        titleCell = firstLine.createCell(2);
        titleCell.setCellValue("供应商名称");
        titleCell = firstLine.createCell(3);
        titleCell.setCellValue("付款（含税）总金额");
        titleCell = firstLine.createCell(4);
        titleCell.setCellValue("发票号");
        titleCell = firstLine.createCell(5);
        titleCell.setCellValue("备注");
        titleCell = firstLine.createCell(6);
        titleCell.setCellValue("登记日期");
        titleCell = firstLine.createCell(7);
        titleCell.setCellValue("银行账号");
        int cellLine = 2;

        //extractor
        ArrayList<WordInfo> orderList = new ArrayList<>();
        WordInfo info;
        for (SelectFile file : mSelectFileList) {
            File excelFile = new File(file.getPath());
            FileInputStream is;
            Workbook workbook;
            if (file.getName().endsWith(".xlsx")) {
                //new excel file
                is = new FileInputStream(excelFile);
                workbook = new XSSFWorkbook(is);
                Sheet sheet = workbook.getSheetAt(0);
                int firstRowNum = sheet.getFirstRowNum();
                int lastRowNum = sheet.getLastRowNum();
                for (int i = 1; i < lastRowNum; i++) {
                    int number = firstRowNum + i;
                    Row row = sheet.getRow(number);
                    //以WrodInfo 暂代Excel台账信息
                    info = new WordInfo();
                    info.setOrderNumber(getCellValue(row.getCell(1)));//订单编号
                    info.setTotalPrice(getCellValue(row.getCell(10)));//总金额
                    info.setSupplyOrderNumber(getCellValue(row.getCell(20)));//发票号

                    orderList.add(info);

                    /*
                    System.out.println("number :" + number + "~~订单编号~~" + getCellValue(row.getCell(1)));
                    System.out.println("number :" + number + "~~总金额~~" + getCellValue(row.getCell(10)));
                    System.out.println("number :" + number + "~~发票号~~" + getCellValue(row.getCell(20)));

                    Row newrow = outSheet.createRow(cellLine);
                    newrow.createCell(0).setCellValue(cellLine - 1);//序号
                    newrow.createCell(1).setCellValue(getCellValue(row.getCell(1)));//采购订单号
                    newrow.createCell(2).setCellValue("");//供应商名称
                    newrow.createCell(3).setCellValue(getCellValue(row.getCell(10)));//总金额
                    newrow.createCell(4).setCellValue(getCellValue(row.getCell(20)));//发票
                    newrow.createCell(5).setCellValue("");//备注
                    newrow.createCell(6).setCellValue("");//日期
                    newrow.createCell(6).setCellValue("");//银行账号
                    cellLine++;*/
                }
            } else {
                //old excel file
            }

            String lastOrderNumber = "";
            for (WordInfo wordInfo : orderList) {
                if (wordInfo.getOrderNumber().length() != 0) {
                    lastOrderNumber = wordInfo.getOrderNumber();
                } else {
                    wordInfo.setOrderNumber(lastOrderNumber);
                }
                System.out.println("orderList Excel Item: "
                        + wordInfo.getOrderNumber() + " | " + wordInfo.getTotalPrice());
            }
            orderList = doSumPrice(orderList);
            for (WordInfo wordInfo : orderList) {
                System.out.println("Excel Item: " + wordInfo.getOrderNumber() + " | " + wordInfo.getTotalPrice());
            }
            for (WordInfo item : orderList) {
                Row newrow = outSheet.createRow(cellLine);
                newrow.createCell(0).setCellValue(cellLine - 1);//序号
                newrow.createCell(1).setCellValue(item.getOrderNumber());//采购订单号
                newrow.createCell(2).setCellValue("");//供应商名称
                newrow.createCell(3).setCellValue(String.format("%.2f",Double.valueOf(item.getTotalPrice())));//总金额
                newrow.createCell(4).setCellValue(item.getSupplyOrderNumber());//发票
                newrow.createCell(5).setCellValue("");//备注
                newrow.createCell(6).setCellValue("");//日期
                newrow.createCell(6).setCellValue("");//银行账号
                cellLine++;
            }

            wirteBook.write(new FileOutputStream(outputExcelFile));
            wirteBook.close();
        }
    }

    private static String getCellValue(Cell cell) {
        String returnValue = "";
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
                returnValue = cell.getNumericCellValue() + "";
                break;
            case STRING:
                returnValue = cell.getRichStringCellValue().getString();
                break;
            case BLANK:
            case _NONE:
            case ERROR:
                returnValue = "";
        }
        return returnValue;
    }

    private static ArrayList<WordInfo> doSumPrice(ArrayList<WordInfo> list) {
        ArrayList<WordInfo> returnList = new ArrayList<>();

        for (WordInfo item : list) {
            if (ifInList(returnList, item)) {
                WordInfo info = getTargetItem(returnList, item);
                if (info == null)
                    break;
                float currentPrice = Float.valueOf(info.getTotalPrice());
                float newPrice = Float.valueOf(item.getTotalPrice());
                info.setTotalPrice((currentPrice + newPrice) + "");
            } else {
                returnList.add(item);
            }
        }
        return returnList;
    }

    private static boolean ifInList(ArrayList<WordInfo> list, WordInfo item) {
        boolean isIn = false;
        for (WordInfo info : list) {
            if (info.getOrderNumber().equals(item.getOrderNumber())) {
                isIn = true;
            }
        }
        return isIn;
    }

    private static WordInfo getTargetItem(ArrayList<WordInfo> list, WordInfo item) {
        for (WordInfo info : list) {
            if (info.getOrderNumber().equals(item.getOrderNumber())) {
                return info;
            }
        }
        return null;
    }
    //</editor-fold>


}
