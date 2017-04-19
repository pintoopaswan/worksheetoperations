package com.workbook.serviceImpl;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.workbook.service.WorkbookService;

public class WorkbookServiceImpl implements WorkbookService {

	static XSSFRow row;

	public XSSFWorkbook createWorkbook(String filename) {

		// create blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		workbook.createSheet();
		try {
			if (filename.isEmpty() || filename.trim().length() == 0) {
				System.out.println("File Name is required");
				return null;
			}

			// create file system using specific name
			if (new File(filename + ".xlsx").isFile()) {
				System.out.println(filename + ".xlsx File already exists!!");
				return workbook;
			}
			FileOutputStream out = new FileOutputStream(new File(filename + ".xlsx"));
			// write operation in workbook using file out object
			workbook.write(out);
			out.close();
			System.out.println(filename + ".xlsx created successfully");
		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
		return workbook;
	}

	public void OpenWorkBook(String filename) {
		try {
			File file = new File(filename + ".xlsx");
			FileInputStream fip = new FileInputStream(file);
			// Get the workbook instance for XLSX file
			XSSFWorkbook workbook = new XSSFWorkbook(fip);
			if (file.isFile() && file.exists()) {
				System.out.println(filename + ".xlsx file open successfully.");
			} else {
				System.out.println("Error to open file.");
			}
		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
	}

	@Override
	public XSSFSheet createSheet(String workbookname, String sheetname) {
		XSSFSheet sheet = null;
		try {
			FileInputStream fis = new FileInputStream(new File(workbookname+".xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			System.out.println("No of sheets before:"+workbook.getNumberOfSheets());
			if (workbook.getSheet(sheetname) != null) {
				System.out.println("spreadsheet already exists!!");
			}else{
				FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
				sheet=workbook.createSheet(sheetname);
				workbook.write(out);
				out.close();
				System.out.println("spreadsheet created");
				System.out.println("No of sheets after:"+workbook.getNumberOfSheets());
			}

		} catch (Exception e) {
			System.out.println("exception:" + e.getMessage());
		}
		return sheet;
	}

	@Override
	public void writeSheet(String workbookname, String sheetname, Map<Object, Object[]> data) {
		System.out.println(" write sheet started successfully");
		try {

			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet(sheetname);
			// Create row object
			XSSFRow row;

			// Iterate over data and write to sheet
			Set<Object> keyid = data.keySet();
			int rowid = 0;
			for (Object key : keyid) {
				row = sheet.createRow(rowid++);
				Object[] objArr = data.get(key);
				int cellid = 0;
				for (Object obj : objArr) {
					Cell cell = row.createCell(cellid++);
					if (obj instanceof String) {

						cell.setCellValue((String) obj);
					} else if (obj instanceof Integer) {
						cell.setCellValue((Integer) obj);

					}
				}
			}
			FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
			workbook.write(out);
			out.close();
			System.out.println(" written successfully");

		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
	}

	public void readSheet(String filename) {
		try {
			System.out.println("Reading: "+ filename+".xlsx");
			FileInputStream fis = new FileInputStream(new File(filename+".xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = spreadsheet.iterator();
			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// System.out.println(cell);

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + " \t\t\t ");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + " \t\t ");
						break;

					case Cell.CELL_TYPE_FORMULA:
						FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
						String value = formulaEval.evaluate(cell).formatAsString();
						System.out.print(value + " \t\t\t ");
						break;
					}
				}
				System.out.println();
			}
			fis.close();
			System.out.println("Closing: "+ filename+".xlsx");
		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
	}

	@Override
	public void useFormula(String workbookname,String sheetname) {
		try {
			System.out.println("Workbook Name:"+workbookname);
			FileInputStream fis = new FileInputStream(new File(workbookname+".xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet(sheetname);
			
			//creating new sheet and writing data into it.
			XSSFSheet sheet1 = workbook.createSheet("sheet1");
			XSSFRow row1=sheet1.createRow(0);
			XSSFCell cell1=row1.createCell(0);
			cell1.setCellValue("sheet2");
			
			workbook.createSheet();

			System.out.println("No of sheets:"+workbook.getNumberOfSheets());
			System.out.println("Hidden:"+workbook.isSheetHidden(1));
			
			
			FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
			
			//setting font color and foreground colors
			 XSSFCellStyle style = workbook.createCellStyle();
			 style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			 style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			 Font font = workbook.createFont();
	         font.setColor(IndexedColors.RED.getIndex());
	         style.setFont(font);
			 
			 
			XSSFRow row = sheet.getRow(0);
			XSSFCell cell = row.createCell(4);
			row.getCell(4).setCellStyle(style);
			cell.setCellValue("Total");
			
			row = sheet.getRow(1);
			cell = row.createCell(4);
			cell.setCellFormula("SUM(B2:C2:D2)");
			
			row = sheet.getRow(2);
			cell = row.createCell(4);
			cell.setCellFormula("SUM(B3:C3:D3)");
			
			XSSFRow row4=sheet.createRow(5);
			XSSFCell cell40 = row4.createCell(0);
			cell40.setCellValue("Hidden Test");
			row4.getCell(0).setCellStyle(style);
			
			//Hiding rows and columns
			row4.getCTRow().setHidden(true);
			sheet.setColumnHidden(4,true);
			//workbook.setSheetHidden(2,true);
			sheet.setColumnHidden(4,false);
			//row4.getCTRow().setHidden(false);
			
			workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
	}

	@Override
	public void usevlookup(String workbookname) {
		try {
			System.out.println("Workbook Name:"+workbookname);
			FileInputStream fis = new FileInputStream(new File(workbookname+".xls"));
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheet("Voters");
			int rowCount=sheet.getLastRowNum();
			FileOutputStream out = new FileOutputStream(new File(workbookname + ".xls"));
			
			//setting font color and foreground colors
			 HSSFCellStyle style = workbook.createCellStyle();
			 style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			 style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			 Font font = workbook.createFont();
	         font.setColor(IndexedColors.RED.getIndex());
	         style.setFont(font);
	         
				HSSFRow row = sheet.getRow(0);
				HSSFCell cell = row.createCell(3);
				row.getCell(3).setCellStyle(style);
				cell.setCellValue("Party Name");
			 
	        System.out.println("No of Rows:"+rowCount); 
	        for(int i=1;i<=rowCount;i++){
	        	row = sheet.getRow(i);
	        	cell = row.createCell(3);
	        	cell.setCellType(Cell.CELL_TYPE_FORMULA);
	        	
	        	//getting cell reference C2,C3....
	        	 String cellRef = new CellReference(row.getRowNum(),
	                     2, false, false).formatAsString();
	        	 System.out.println("cell ref:"+cellRef);
	        	 String strFormula = "VLOOKUP("+cellRef+",'Party Codes'!A2:B45,2,False)";
	        	cell.setCellFormula(strFormula);
	        }
	        
			workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
			workbook.write(out);
			out.close();
			System.out.println(" Vlookup completed successfully");
		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
		
	}
	
	@Override
	public void writeInChart(String workbookname, String sheetname, Map<Object, Object[]> data) {
		System.out.println(" write In Chart started successfully");
		try {

			FileInputStream fis = new FileInputStream(new File(workbookname+".xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet(sheetname);
			// Create row object
			XSSFRow row;

			// Iterate over data and write to sheet
			Set<Object> keyid = data.keySet();
			int rowid = 0;
			for (Object key : keyid) {
				row = sheet.createRow(rowid++);
				Object[] objArr = data.get(key);
				int cellid = 0;
				for (Object obj : objArr) {
					Cell cell = row.createCell(cellid++);
					if (obj instanceof String) {

						cell.setCellValue((String) obj);
					} else if (obj instanceof Integer) {
						cell.setCellValue((Integer) obj);

					}
				}
			}
			FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
			workbook.write(out);
			out.close();
			System.out.println(" written successfully");

		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
	}

	@Override
	public void renameSheet(String workbookname, String oldSheetname, String newSheetname) {
		System.out.println(" Rename Sheet started successfully");
		try {
			FileInputStream fis = new FileInputStream(new File(workbookname+".xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			
			if(workbook.getSheet(newSheetname)!=null){
				System.out.println("New sheet name already exists!!");
				
			}else{
				int index=workbook.getSheetIndex(oldSheetname);
				workbook.setSheetName(index,newSheetname);
				FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
				workbook.write(out);
				out.close();
				System.out.println("sheet renamed!!");
			}
			
			
		}catch(IllegalArgumentException ie){
			System.out.println("Invalid  sheet name");
		}
		catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	@Override
	public void protctedWorkbook(String workbookname) {
		System.out.println("protctedWorkbook");
        BufferedInputStream bufferInput = null;      
        POIFSFileSystem poiFileSystem = null;    
        FileOutputStream fileOut = null;
        try {    
        	String fname = workbookname+".xls";
    		FileInputStream fileInput = new FileInputStream(new File(fname));
    		    

         
            bufferInput = new BufferedInputStream(fileInput);            
            poiFileSystem = new POIFSFileSystem(bufferInput);            

            Biff8EncryptionKey.setCurrentUserPassword("secret");            
            HSSFWorkbook workbook = new HSSFWorkbook(poiFileSystem, true);            
            HSSFSheet sheet = workbook.getSheetAt(0);           

            HSSFRow row = sheet.createRow(0);
            Cell cell = row.createCell(0);

            cell.setCellValue("THIS WORKS!"); 

            fileOut = new FileOutputStream(fname);
            workbook.writeProtectWorkbook(Biff8EncryptionKey.getCurrentUserPassword(), "");
            workbook.write(fileOut);



        } catch (Exception ex) {

            System.out.println(ex.getMessage());      

        } finally {         

              try {            

                  bufferInput.close();     

              } catch (IOException ex) {

                  System.out.println(ex.getMessage());     

              }    

              try {            

                  fileOut.close();     

              } catch (IOException ex) {

                  System.out.println(ex.getMessage());     

              } 
		
	}
        }

	@Override
	public void formatSheet(String workbookname, String sheetname, Map<Object, Object[]> data) {
		System.out.println(" write sheet started successfully");
		try {
			
//			FileInputStream fis = new FileInputStream(new File(workbookname+".xlsx"));
//			XSSFWorkbook workbook = new XSSFWorkbook(fis);

			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet(sheetname);
			// Create row object
			XSSFRow row;

			// Iterate over data and write to sheet
			Set<Object> keyid = data.keySet();
			int rowid = 0;
			for (Object key : keyid) {
				row = sheet.createRow(rowid++);
				Object[] objArr = data.get(key);
				
				int cellid = 0;
				for (Object obj : objArr) {
					
					if(obj.toString().equals("Marks Details")){
						Cell cell = row.createCell(1);
						cell.setCellValue((String) obj);
						sheet.addMergedRegion(new CellRangeAddress(0,0,1,4));
					}
					else if(obj.toString().equals("1st Semester")){
						Cell cell = row.createCell(1);
						cell.setCellValue((String) obj);
						sheet.addMergedRegion(new CellRangeAddress(1,1,1,2));
					}
					else if(obj.toString().equals("2nd Semester")){
						Cell cell = row.createCell(3);
						cell.setCellValue((String) obj);
						sheet.addMergedRegion(new CellRangeAddress(1,1,3,4));
					}else{
					Cell cell = row.createCell(cellid++);
					if (obj instanceof String) {

						cell.setCellValue((String) obj);
						
					} else if (obj instanceof Integer) {
						cell.setCellValue((Integer) obj);

					}
					}
				}
			}
			FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
			workbook.write(out);
			out.close();
			System.out.println(" written successfully");

		} catch (Exception e) {
			System.out.println(e.getStackTrace());
		}
		
	}

	@Override
	public void readColumn(String workbookname, String sheetname, int colIndexToRead) {
		System.out.println("Reading column strated...");
		try{
		FileInputStream fis = new FileInputStream(new File(workbookname+".xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheet(sheetname);
		for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			  row = sheet.getRow(rowIndex);
			  if (row != null) {
			   // String cellValueMay = null;
			    for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
			      if (colIndex == colIndexToRead) {
			        XSSFCell cell = row.getCell(colIndex);
			        if (cell != null) {
			          // Found column and there is value in the cell.
			          System.out.println(cell.getStringCellValue()); 
			          break;
			        }
			    }

			    // Do something with the cellValueMaybeNull here ...
			  }
			}
		}
		System.out.println("Reading column End...");
		}catch(Exception e){
			System.out.println("Exception Occured:"+e.getMessage());
		}
	}
	
@Override
 public void sortSheet(String workbookname,String sheet1, int column, int rowStart) {
	 boolean sorting = true;
	 try{
		FileInputStream fis = new FileInputStream(new File(workbookname));
		HSSFWorkbook workbook = new HSSFWorkbook(fis);
		HSSFSheet sheet=workbook.getSheet(sheet1);
	    int lastRow = sheet.getLastRowNum();
	    while (sorting == true) {
	        sorting = false;
	        for (Row row : sheet) {
	            // skip if this row is before first to sort
	            if (row.getRowNum()<rowStart) continue;
	            // end if this is last row
	            if (lastRow==row.getRowNum()) break;
	            Row row2 = sheet.getRow(row.getRowNum()+1);
	            if (row2 == null) continue;
	            System.out.println(row.getCell(column));
	            System.out.println(row2.getCell(column));
	            String firstValue = (row.getCell(column) != null) ? row.getCell(column).getStringCellValue() : "";
	            String secondValue = (row2.getCell(column) != null) ? row2.getCell(column).getStringCellValue() : "";
	            //compare cell from current row and next row - and switch if secondValue should be before first
	            if (secondValue.compareToIgnoreCase(firstValue)<0) {                          
	                sheet.shiftRows(row2.getRowNum(), row2.getRowNum(), -1);
	                sheet.shiftRows(row.getRowNum(), row.getRowNum(), 1);
	                sorting = true;
	            }
	        }
	    }
	    System.out.println("sorting done");
	 }catch(Exception e){
		 System.out.println("Exception Occured:"+e.getMessage());
	 }
	}

public void utilSheet(String workbookname,String sheet1) {
	 try{
		FileInputStream fis = new FileInputStream(new File(workbookname));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheet(sheet1);
	    int lastRow = sheet.getLastRowNum();
	    System.out.println("rows:"+lastRow);
	    
	 }catch(Exception e){
		 System.out.println("Exception Occured:"+e.getMessage());
	 }
}
}
