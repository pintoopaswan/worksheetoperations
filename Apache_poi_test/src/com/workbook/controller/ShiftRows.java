package com.workbook.controller;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ShiftRows {
	 public static void main(String[]args) throws IOException {
	        Workbook wb = new XSSFWorkbook();   //or new HSSFWorkbook();
	        Sheet sheet = wb.createSheet("Sheet1");

	        Row row1 = sheet.createRow(2);
	        row1.createCell(0).setCellValue(1);

	        Row row2 = sheet.createRow(3);
	        row2.createCell(0).setCellValue(2);

	        Row row3 = sheet.createRow(4);
	        row3.createCell(0).setCellValue(3);

	        Row row4 = sheet.createRow(5);
	        row4.createCell(0).setCellValue(4);

	        Row row5 = sheet.createRow(6);
	        row5.createCell(0).setCellValue(5);
	        
	        row1.getCell(0).setCellValue(8);

	        //Shift rows 6 - 11 on the spreadsheet to the top (rows 0 - 5)
	        //sheet.shiftRows(2, 2, 1);
	        //sheet.shiftRows(3, 3, -1);

	        FileOutputStream fileOut = new FileOutputStream("shiftRows.xlsx");
	        wb.write(fileOut);
	        fileOut.close();
	        wb.close();
	        System.out.println("Shifting Done");
	    }


}
