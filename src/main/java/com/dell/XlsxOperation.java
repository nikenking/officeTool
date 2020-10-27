package com.dell;

import com.dell.util.Tool;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class XlsxOperation {
    @Test
    public void xlsxtest1() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("FirstTable");
        Row r1 = sheet.createRow(0);
        Cell c1 = r1.createCell(0);
        c1.setCellValue("第一行第一列的数据");
        Cell c2 = r1.createCell(1);
        c2.setCellValue(123);

        Row r2 = sheet.createRow(1);
        Cell c21 = r2.createCell(0);
        c21.setCellValue("第二行第一列的数据,右边即将显示当前时间");
        Cell c22 = r2.createCell(1);
        c22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        String path = "D:\\project\\maven2\\";
        FileOutputStream fo = new FileOutputStream(path+"创造101.xlsx");
        workbook.write(fo);
        fo.close();
        System.out.println("第一张表👌了");
    }
    @Test
    public void xlsxtest2() throws IOException{
        long start = System.currentTimeMillis();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("sheet1");
        for (int row = 0; row < 65536; row++) {
            Row row1 = sheet1.createRow(row);
            for (int col = 0; col < 10; col++) {
                Cell cell = row1.createCell(col);
                cell.setCellValue(col);
            }
        }
        String path = "D:\\project\\maven2\\";
        FileOutputStream fo = new FileOutputStream(path+"bigOne.xlsx");
        workbook.write(fo);
        fo.close();
        long end = System.currentTimeMillis();
        System.out.println((double)(end-start));
    }
    @Test
    public void xlsxtest3() throws IOException{
        String head = "第十一周";
        Tool.rename(head);
        String path = "C:\\Users\\15756\\Desktop\\周总结\\";
        File file = new File(path);
        List<List<String>> lists = new ArrayList<>();
        List<File> files = Tool.fileSort(file.listFiles());
        for (File listFile : files) {
            Matcher m = Pattern.compile(".*?(.{1,2}组).*?").matcher(listFile.toString());
            while (m.find()) {//是一个可读xslx文件
                List<String> tmp = new ArrayList<>();
                tmp.add(m.group(1));
                lists.add(tmp);
                FileInputStream fi = new FileInputStream(listFile);//输入流转换
                Workbook workbook = new XSSFWorkbook(fi);
                Sheet sheet = workbook.getSheetAt(0);
                int rowCount = sheet.getPhysicalNumberOfRows();//所有行
                for (int i = 1; i < rowCount; i++) {//除了第一行
                    Row row = sheet.getRow(i);//第二行开始,遍历每一个人
                    if (row!=null){
                        List<String> list = new ArrayList<>();
                        int cells = row.getPhysicalNumberOfCells();//一个人
                        for (int cellNum = 0; cellNum < cells; cellNum++) {//开始遍历个人数据
                            Cell cell = row.getCell(cellNum);//个人具体值
                            if (cell!=null&&cell.getCellType()==1){
                                if (Pattern.matches(".*?[\\u4e00-\\u9fa5|\\w].*?",cell.getStringCellValue())) {//真人,不是空白填充
                                    String value = cell.getStringCellValue();
                                    list.add(value);
                                }
                            }
                        }
                        if (list.size()>0){
                            lists.add(list);
                        }
                    }
                }
                lists.add(new ArrayList<String>());
                fi.close();
            }
        }
        for (List<String> list : lists) {
            for (String s : list) {
                System.out.print(s+"-----:-----");
            }
            System.out.println();
        }
        Tool.insert(lists,head);
    }
}
