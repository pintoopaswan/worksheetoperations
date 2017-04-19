package com.workbook.service;

import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface WorkbookService {
	public XSSFWorkbook createWorkbook(String workbookname);

	public void OpenWorkBook(String workbookname);

	public XSSFSheet createSheet(String workbookname, String sheetname);
	
	public void readSheet(String sheetname);
	
	void writeSheet(String workbookname, String sheetname, Map<Object, Object[]> data);
	
	public void renameSheet(String workbookname,String oldSheetname,String newSheetname);

	public void useFormula(String workbookname,String sheetname);

	void usevlookup(String workbookname);
	
	void writeInChart(String workbookname, String sheetname, Map<Object, Object[]> data);
	
	void protctedWorkbook(String workbookname);
	
	void formatSheet(String workbookname, String sheetname, Map<Object, Object[]> data);
	
	void readColumn(String workbookname, String sheetname,int colIndexToRead);
	
	 void sortSheet(String workbookname, String sheetname, int column, int rowStart);

}
