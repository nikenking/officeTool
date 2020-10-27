package com.dell.util;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.xmlbeans.XmlException;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Tool {
    public static void rename(String head) {
        String path = "C:\\Users\\15756\\Desktop\\周总结\\";
        System.out.println("当前查找目录是:" + path);
//        String head = new Scanner(System.in).nextLine();
        File file = new File(path);
        for (File listFile : file.listFiles()) {
            Matcher m = Pattern.compile("[\\u4e00-\\u9fa5]{1,2}(?=组)").matcher(listFile.toString());
            while (m.find()) {
                boolean b = listFile.renameTo(new File(path + head + "-第" + m.group() + "组-进度汇报.xlsx"));
                System.out.println(listFile.toString());
            }
        }
    }

    public static void insert(List<List<String>> lists, String head) throws IOException {
        String path = "C:\\Users\\15756\\Desktop\\周总结\\" + head + "-进度汇报汇总.xlsx";
        FileOutputStream fo = new FileOutputStream(path);
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("第十周汇总");
        for (int i = 0; i < lists.size(); i++) {
            Row row = sheet.createRow(i);//每个人
            List<String> list = lists.get(i);
            for (int i1 = 0; i1 < list.size(); i1++) {
                Cell cell = row.createCell(i1);//每个人的具体数据
                cell.setCellValue(list.get(i1));
            }
        }
        workbook.write(fo);
        fo.close();
    }

    public static List<File> fileSort(File[] files) {
        List<File> list = new ArrayList<>();
        Map<Integer, String> arr = new HashMap<>();
        for (int i = 0; i < files.length; i++) {
            if (Pattern.matches(".*?一组.*?", files[i].toString())) {
                arr.put(0, files[i].toString());
            } else if (Pattern.matches(".*?二组.*?", files[i].toString())) {
                arr.put(1, files[i].toString());
            } else if (Pattern.matches(".*?三组.*?", files[i].toString())) {
                arr.put(2, files[i].toString());
            } else if (Pattern.matches(".*?四组.*?", files[i].toString())) {
                arr.put(3, files[i].toString());
            } else if (Pattern.matches(".*?五组.*?", files[i].toString())) {
                arr.put(4, files[i].toString());
            } else if (Pattern.matches(".*?六组.*?", files[i].toString())) {
                arr.put(5, files[i].toString());
            } else if (Pattern.matches(".*?七组.*?", files[i].toString())) {
                arr.put(6, files[i].toString());
            } else if (Pattern.matches(".*?八组.*?", files[i].toString())) {
                arr.put(7, files[i].toString());
            } else if (Pattern.matches(".*?九组.*?", files[i].toString())) {
                arr.put(8, files[i].toString());
            } else if (Pattern.matches(".*?十组.*?", files[i].toString())) {
                arr.put(9, files[i].toString());
            }
        }
        for (Map.Entry<Integer, String> e : arr.entrySet()) {
            list.add(new File(e.getValue()));
        }
        return list;
    }

    /*word判断*/
    public static void findKeyWords(String path, String msg) throws IOException, OpenXML4JException, XmlException {
        OPCPackage opcPackage = POIXMLDocument.openPackage(path);
        XWPFWordExtractor extractor = new XWPFWordExtractor(opcPackage);
        String text = extractor.getText();
        if (text.contains(msg)) {
            System.out.println(path);
        }
    }

    /*写入list到excel*/
    public static void listInsertExcel(List<String> list) throws IOException {
        String path = "C:\\Users\\15756\\Desktop\\一班小组成员.xlsx";
        FileOutputStream fo = new FileOutputStream(path);
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("一班分组");
        int id = 1;
        for (int i = 1; i < list.size() + 1; i++) {
            Row row = sheet.createRow(i);//一个人
            Cell cell = row.createCell(0);
            cell.setCellValue(id);
            Cell cell1 = row.createCell(1);
            cell1.setCellValue(list.get(i - 1).split(":")[0]);
            Cell cell2 = row.createCell(2);
            cell2.setCellValue(list.get(i - 1).split(":")[1]);
            id++;
        }
        workbook.write(fo);
        fo.close();
    }
}
