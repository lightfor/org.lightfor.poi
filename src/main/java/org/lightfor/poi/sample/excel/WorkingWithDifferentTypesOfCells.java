package org.lightfor.poi.sample.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

/**
 * WorkingWithDifferentTypesOfCells
 * Created by Light on 2016/12/22.
 */
public class WorkingWithDifferentTypesOfCells {

    public static void main(String[] args) throws Exception{
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        Row row = sheet.createRow((short)2);
        row.createCell(0).setCellValue(1.1);
        row.createCell(1).setCellValue(new Date());
        row.createCell(2).setCellValue(Calendar.getInstance());
        row.createCell(3).setCellValue("a string");
        row.createCell(4).setCellValue(true);
        //row.createCell(5).setCellType(CellType.ERROR);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
