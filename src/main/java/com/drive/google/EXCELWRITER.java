package com.drive.google;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.time.LocalDateTime.*;
import java.io.FileOutputStream;
import java.io.IOException;

public class EXCELWRITER {

    public static void main(String args[]) {

        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("ExcelSheet");

        Object[][] softwareVersion = {{"Apache POI", "Version", 2}, {"TestNG", "Version", 1},
                {"Extent report", "Version", 4}, {"Selenium", "Version", 4}
        };
        int rowNumber = 0;

        for (Object[] cRow : softwareVersion) {

            Row row = sheet.createRow(++rowNumber);
            int cellNumber = 0;

            for (Object cCell : cRow) {
                Cell cell = row.createCell(++cellNumber);
                if (cCell instanceof String) {
                    cell.setCellValue((String) cCell);
                } else if (cCell instanceof Integer) {
                    cell.setCellValue((Integer) cCell);
                }
            }
        }

        try (FileOutputStream fOut = new FileOutputStream("ExcelSheet.xlsx")) {
            wb.write(fOut);

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
