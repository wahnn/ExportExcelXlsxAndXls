package com.sl.excelXls;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Illustrates how to create cell values.
 *
 * @author Glen Stampoultzis (glens at apache.org)
 */
public class CreateCells {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFSheet sheet = wb.createSheet("new sheet");

            // Create a row and put some cells in it. Rows are 0 based.
            HSSFRow row = sheet.createRow(0);
            // Create a cell and put a value in it.
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(1);

            // Or do it on one line.
            row.createCell(1).setCellValue(1.2);
            row.createCell(2).setCellValue("This is a string");
            row.createCell(3).setCellValue(true);

            String path = "D:\\excelExport";
            File file = new File(path);
            if(!file.exists()){
                file.mkdirs();
            }

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream(path+File.separator+"workbook.xls")) {
                wb.write(fileOut);
            }
        }
    }
}