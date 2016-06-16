package com.util;

/**
 * Created by cyb on 2016/6/16.
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtil {

    public static void main(String[] args) {
        File file = new File("/");
        String path2007 = file.getAbsolutePath() + "excel\\src\\main\\resources\\test.xlsx";
        parseExcel(path2007);
    }
    /**
     * 根据路径加载解析Excel
     *
     * @param path
     * @return
     */
    public static String parseExcel(String path) {
        File file = null;
        InputStream input = null;
        Workbook workBook = null;
        Sheet sheet = null;
        file = new File(path);
        try {
            input = new FileInputStream(file);
            workBook = WorkbookFactory.create(input);
            if (workBook != null) {
                int numberSheet = workBook.getNumberOfSheets();
                if (numberSheet > 0) {
                    sheet = workBook.getSheetAt(0);//获取第一个工作簿(Sheet)的内容
                    getExcelContent(sheet);
                } else {
                    System.out.println("目标表格工作簿(Sheet)数目为0！");
                }
            }
            input.close();
        } catch (Exception e) {
            System.out.println("关闭输入流异常！" + e.getMessage());
            e.printStackTrace();
        }

        return "";
    }

    @SuppressWarnings("static-access")
    public static String getExcelContent(Sheet sheet) {
        int rowCount = sheet.getPhysicalNumberOfRows();//总行数
        int k = 0;
        if (rowCount > 1) {
            for (int i = 0; i < rowCount; i++) {//遍历行，略过标题行，从第二行开始
                Row row = sheet.getRow(i);
                for (int j = 0; j < 1; j++) {
                    Cell cell = row.getCell(j);
                    String bd = cell.getStringCellValue();
                    System.out.println("insert into t_cj_promo_code values(" + ++k + ",'" + bd + "',0,sysdate,8); ");
                }
            }
        }
        return "";
    }
}
