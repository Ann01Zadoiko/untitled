package org.example;

import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * Hello world!
 *
 */
public class App {

    static XSSFRow row;
    public static void main(String[] args ) throws IOException {

// FileInputStream fileInputStream = new FileInputStream(".\\src\\main\\resources\\NumberOfTrams.xlsx")) {


//        try {
//            File file = new File(".\\src\\main\\resources\\NumberOfTrams.xlsx");
//            FileInputStream fIP = new FileInputStream(file);
//
//            //Get the workbook instance for XLSX file
//            XSSFWorkbook workbook = new XSSFWorkbook(fIP);
//
//            if(file.isFile() && file.exists()) {
//                System.out.println("openworkbook.xlsx file open successfully.");
//            } else {
//                System.out.println("Error to open openworkbook.xlsx file.");
//            }
//        } catch(Exception e) {
//            System.out.println("Error to open openworkbook.xlsx file." + e.getMessage());
//        }

        FileInputStream fis = new FileInputStream(new File(".\\src\\main\\resources\\NumberOfTrams.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet spreadsheet = workbook.getSheetAt(0);
        Iterator < Row >  rowIterator = spreadsheet.iterator();

        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();
            Iterator < Cell >  cellIterator = row.cellIterator();

            while ( cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getCellType()) {
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + " \t\t ");
                        break;

                    case STRING:
                        System.out.print(
                                cell.getStringCellValue() + " \t\t ");
                        break;
                }
            }
            System.out.println();
        }
        fis.close();
    }
}
