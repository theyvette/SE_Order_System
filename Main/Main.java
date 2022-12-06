import java.util.Scanner;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
 
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;

public class Main {
	private static String getTimeNumber() {
        String pattern = "yyyyMMddHHmmssSSS";
        SimpleDateFormat d = new SimpleDateFormat(pattern);
        return d.format(new Date());
    }
     
    public static void main(String[] args) {
                 
        @SuppressWarnings("resource")
        // 新建工作簿
        XSSFWorkbook book = new XSSFWorkbook();
        // 建立工作表
        XSSFSheet sheet = book.createSheet("Books");
 
        Object[][] buffer = { 
                { "Head First Java", "Kathy Serria", 79 }, 
                { "Effective Java", "Joshua Bloch", 36 },
                { "Clean Code", "Robert martin", 42 }, 
                { "Thinking in Java", "Bruce Eckel", 35 }, 
                };
 
        int rowIdx = -1;
        int colIdx = -1;
         
        CellRangeAddress cellAddr;
        int firstRow, lastRow, firstCol, lastCol;
         
        XSSFRow row;
        XSSFCell cell;
        for (Object[] arrs : buffer) {
            // 建立行
            row = sheet.createRow(++rowIdx);
            firstRow = lastRow = rowIdx;
 
            colIdx = -1;
            firstCol = (colIdx + 1);
            for (Object field : arrs) {
                // 建立單元格
                cell = row.createCell(++colIdx);
                 
                // 單元格寫入內容
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
            lastCol = colIdx;
             
            // BorderStyle.THICK 粗邊框
            cellAddr = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            RegionUtil.setBorderBottom(BorderStyle.THICK, cellAddr, sheet);
        }
 
        // 指定檔案名稱
        String fileName = "JavaBooks_%1$s.xlsx";
        fileName = String.format(fileName, getTimeNumber());
         
        /*
         * 尚未指定檔案路徑，檔案建立在本執行專案內
         * 儲存工作簿
         * */
        try (FileOutputStream os = new FileOutputStream(fileName)) {
            book.write(os);
            System.out.println(fileName + " excel export finish.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}