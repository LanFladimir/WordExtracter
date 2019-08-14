package utils;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import bean.WordInfo;

public class WordReader {


    /**
     * 抽取Word内容
     */
    public static ArrayList<WordInfo> readWord(String wordPath) {
        String orderNumber = "";
        String supplyOrderNumber = "";
        ArrayList<WordInfo> wordInfos = new ArrayList<>();
        WordInfo wordInfo;
        try {
            File docxFile = new File(wordPath);
            FileInputStream is = new FileInputStream(docxFile);
            //此方法读取内容，纯文本
            /*if (wordPath.endsWith(".doc")) {
                WordExtractor ex = new WordExtractor(is);
                System.out.println(ex.getText());
            } else if (wordPath.endsWith(".docx")) {
                OPCPackage opcPackage = POIXMLDocument.openPackage(wordPath);
                POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
                System.out.println(extractor.getText());
            } else {
                System.out.println("文件类型异常：" + wordPath);
            }*/
            //jar:poi-ooxml-3.17.jar
            if (wordPath.toLowerCase().endsWith("docx")) {
                XWPFDocument xwpf = new XWPFDocument(is);
                Iterator<XWPFTable> it = xwpf.getTablesIterator();
                int tableNumber = 0;
                while (it.hasNext() && tableNumber < 2) {
                    XWPFTable table = it.next();
                    List<XWPFTableRow> rows = table.getRows();
                    if (tableNumber == 0) {
                        List<XWPFTableCell> cell01 = rows.get(0).getTableCells();
                        orderNumber = cell01.get(1).getText().split("：")[1];//订单编号
                        List<XWPFTableCell> cell02 = rows.get(1).getTableCells();
                        supplyOrderNumber = cell02.get(1).getText().split("：")[1];//供货单号
                    } else if (tableNumber == 1) {
                        for (int i = 1; i < rows.size() - 1; i++) {
                            List<XWPFTableCell> cell = rows.get(i).getTableCells();

                            wordInfo = new WordInfo();
                            wordInfo.setSrcfile(ExcelWriter.getFileNameWithoutEnd(wordPath));
                            wordInfo.setId(0);
                            wordInfo.setOrderNumber(orderNumber);
                            wordInfo.setSupplyOrderNumber(supplyOrderNumber);
                            wordInfo.setPowerSupply(cell.get(1).getText());
                            wordInfo.setProjectName(cell.get(2).getText());
                            wordInfo.setGoodsName(cell.get(3).getText());
                            wordInfo.setModelName(cell.get(4).getText());
                            wordInfo.setGoodsCloums(cell.get(5).getText());
                            wordInfo.setGoodsUnit(cell.get(6).getText());
                            wordInfo.setNotwithPrice(cell.get(7).getText());
                            wordInfo.setNotwithTotalPrice(cell.get(8).getText());
                            wordInfo.setTotalPrice(cell.get(9).getText());
                            wordInfo.setDeliveryTime(cell.get(10).getText());

                            wordInfos.add(wordInfo);
                        }

                    } else {
                        System.out.println("不需要的表");
                    }

                    tableNumber++;
                }
            } else {
                // 处理doc格式 即office2003版本
                //jar:poi-scratchpad-3.17.jar
                POIFSFileSystem pfs = new POIFSFileSystem(is);
                HWPFDocument hwpf = new HWPFDocument(pfs);
                Range range = hwpf.getRange();//得到文档的读取范围
                TableIterator it = new TableIterator(range);
                // 迭代文档中的表格
                // 如果有多个表格只读取需要的一个 set是设置需要读取的第几个表格，total是文件中表格的总数
                int set = 1, total = 4;
                int num = set;
                for (int i = 0; i < set - 1; i++) {
                    it.hasNext();
                    it.next();
                }
                while (it.hasNext()) {
                    Table tb = (Table) it.next();
                    System.out.println("这是第" + num + "个表的数据");
                    //迭代行，默认从0开始,可以依据需要设置i的值,改变起始行数，也可设置读取到那行，只需修改循环的判断条件即可
                    for (int i = 0; i < tb.numRows(); i++) {
                        TableRow tr = tb.getRow(i);
                        //迭代列，默认从0开始
                        for (int j = 0; j < tr.numCells(); j++) {
                            TableCell td = tr.getCell(j);//取得单元格
                            //取得单元格的内容
                            for (int k = 0; k < td.numParagraphs(); k++) {
                                Paragraph para = td.getParagraph(k);
                                String s = para.text();
                                //去除后面的特殊符号
                                if (null != s && !"".equals(s)) {
                                    s = s.substring(0, s.length() - 1);
                                }
                                System.out.print(s + "\t");
                            }
                        }
                        System.out.println();
                    }
                    // 过滤多余的表格
                    while (num < total) {
                        it.hasNext();
                        it.next();
                        num += 1;
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return wordInfos;
    }
}
