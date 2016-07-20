package com.util;

/**
 * Created by cyb on 2016/6/16.
 */

import java.io.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtil {

    public static void main(String[] args) {
        File file = new File("/");
        String path2007 = file.getAbsolutePath() + "tools\\excelUtil\\src\\main\\resources\\car2.xlsx";
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
                    sheet = workBook.getSheetAt(1);//获取第一个工作簿(Sheet)的内容
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
            String sql = "";
            for (int i = 1; i < rowCount; i++) {//遍历行，略过标题行，从第二行开始
                Row row = sheet.getRow(i);

                String seriesid = null;
                String c1 = null;
                String c2 = null;
                String c3 = null;
                String c4 = null;
                String c5 = null;
                String c6 = null;
                String c7 = null;
                String c8 = null;
                String c9 = null;
                String c10 = null;
                Cell cell = row.getCell(0);
                if(cell != null){
                      seriesid = cell.getStringCellValue();
                }
                Cell cell1 = row.getCell(1);
                if(cell1 != null ){
                     c1 = cell1.getStringCellValue();
                }

                Cell cell2 = row.getCell(2);

                if(cell2 != null ){
                     c2 = cell2.getStringCellValue();
                }

                Cell cell3 = row.getCell(3);
                if(cell3 != null ){
                      c3 = cell3.getStringCellValue();
                }

                Cell cell4 = row.getCell(4);
                if(cell4 != null ){
                      c4 = cell4.getStringCellValue();
                }

                Cell cell5 = row.getCell(5);
                if(cell5 != null ){
                      c5 = cell5.getStringCellValue();
                }

                Cell cell6 = row.getCell(6);
                if(cell6 != null ){
                     c6 = cell6.getStringCellValue();
                }

                Cell cell7 = row.getCell(7);
                if(cell7 != null ){
                      c7 = cell7.getStringCellValue();
                }

                Cell cell8 = row.getCell(8);
                if(cell8 != null ){
                     c8 = cell8.getStringCellValue();
                }

                Cell cell9 = row.getCell(9);
                if(cell9 != null ){
                    c9 = cell9.getStringCellValue();
                }
                Cell cell10 = row.getCell(10);
                if(cell10 != null ){
                      c10 = cell10.getStringCellValue();
                }

               /* System.out.println("insert into compete_series(id,seriesid,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10) " +
                        "values("+i+","+seriesid+","+c1+","+c2+","+c3+","+c4+","+c5+","+c6+","+c7+","+c8+","+c9+","+c10+"); ");
*/
                  sql += "insert into compete_spec(id,specid,s1,s2,s3,s4,s5,s6,s7,s8,s9,s10)values("+i+","+seriesid+","+c1+","+c2+","+c3+","+c4+","+c5+","+c6+","+c7+","+c8+","+c9+","+c10+"); \n ";
            }
            //System.out.println(sql);
            method1(sql);

        }
        return "";
    }


    public static void  method1(String content) {
        FileWriter fw = null;
        try {
            File f=new File("E:\\sql.sql");
            fw = new FileWriter(f, true);
        } catch (IOException e) {
            e.printStackTrace();
        }
        PrintWriter pw = new PrintWriter(fw);
        pw.println(content);
        pw.flush();
        try {
            fw.flush();
            pw.close();
            fw.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void method2(String file, String conent) {
        BufferedWriter out = null;
        try {
            out = new BufferedWriter(new OutputStreamWriter(
                    new FileOutputStream(file, true)));
            out.write(conent+"\r\n");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
