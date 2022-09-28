package com.ocean.excelutils;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * @author huhaiyang
 * @date 2022/9/27
 */
public class ExcelUtils {
    private static final String XLS = ".xls";
    private static final String XLSX = ".xlsx";

    public static List<List<String>> readFromFile(File file) throws IOException {
        List<List<String>> res = new ArrayList<>();
        InputStream in = null;
        try {
            in = new FileInputStream(file);
            String fileName = file.getName();
            String suffix = fileName.substring(fileName.lastIndexOf("."));
            DataFormatter formatter = new DataFormatter(Locale.US);
            Workbook workbook;
            //构造工作簿对象
            switch (suffix) {
                case XLS:
                    workbook = new HSSFWorkbook(in);
                    break;
                case XLSX:
                    workbook = new XSSFWorkbook(in);
                    break;
                default:
                    workbook = new HSSFWorkbook(in);
                    break;
            }
            //获取工作表，默认获取第1个sheet
            Sheet sheet = workbook.getSheetAt(0);
            //获取总行数
            int rowNum = sheet.getLastRowNum();
            //获取总列数
            int columnNum = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int i = 1; i <= rowNum; i++) {
                Row row = sheet.getRow(i);
                List<String> rowList = new ArrayList<>();
                for (int j = 0; j < columnNum; j++) {
                    rowList.add(formatter.formatCellValue(row.getCell(j)));
                }
                res.add(rowList);
            }
        } finally {
            if (in != null) {
                in.close();
            }
        }
        return res;
    }

    public static void writeToFile(List<String> attributes, List<List<String>> data, File file) throws IOException {
        OutputStream out = null;
        try {
            out = new FileOutputStream(file);
            String fileName = file.getName();
            String suffix = fileName.substring(fileName.lastIndexOf("."));
            Workbook workbook;
            //新建 Excel工作簿对象
            switch (suffix) {
                case XLS:
                    workbook = new HSSFWorkbook();
                    break;
                case XLSX:
                    workbook = new XSSFWorkbook();
                    break;
                default:
                    workbook = new HSSFWorkbook();
                    break;
            }
            //新建工作表
            Sheet sheet = workbook.createSheet();
            //建立表格的行
            Row row0 = sheet.createRow(0);
            for (int i = 0; i < attributes.size(); i++) {
                row0.createCell(i).setCellValue(attributes.get(i));
            }
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i + 1);
                for (int j = 0; j < data.get(i).size(); j++) {
                    row.createCell(j).setCellValue(data.get(i).get(j));
                }
            }
            workbook.write(out);
        } finally {
            if (out != null) {
                out.close();
            }
        }
    }


}
