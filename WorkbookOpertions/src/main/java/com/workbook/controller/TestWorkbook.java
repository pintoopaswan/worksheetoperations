package com.workbook.controller;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.workbook.serviceImpl.WorkbookServiceImpl;
public class TestWorkbook {
	

	public static void main(String[] args) throws IOException {
		
		WorkbookServiceImpl WorkbookServiceImpl=new WorkbookServiceImpl();
	
		Map<Object,Object[]> details=new TreeMap<Object,Object[]>();
		details.put("1",new Object[]{"Name","Physics"," Chemistry","Maths","English"});
		details.put("2",new Object[]{"Pintoo",60,50,80,60});
		details.put("3",new Object[]{"Ram",40,80,85,75});
		details.put("4",new Object[]{"Nilima",70,60,95,85});
		
		Map<Object,Object[]> semester=new TreeMap<Object,Object[]>();
		semester.put("1",new Object[]{"Marks Details"});
		semester.put("2",new Object[]{"1st Semester","2nd Semester"});
		semester.put("3",new Object[]{"Student Name","Physics"," Chemistry","Maths","English"});
		semester.put("4",new Object[]{"Pintoo",60,50,80,60});
		semester.put("5",new Object[]{"Ram",40,80,85,75});
		semester.put("6",new Object[]{"Nilima",70,60,95,85});
		
		
		
		XSSFWorkbook workbook=WorkbookServiceImpl.createWorkbook("Student");
		//WorkbookServiceImpl.createSheet("Student", "aaa");
		//WorkbookServiceImpl.writeSheet("Student","marksheet",details);
		//WorkbookServiceImpl.useFormula( "Student","marksheet");
		//WorkbookServiceImpl.readSheet("Student");
		//WorkbookServiceImpl.renameSheet("Student", "Sheet2", "Pintoo1");
		//WorkbookServiceImpl.usevlookup("voters");
		//WorkbookServiceImpl.writeInChart("chart","Sheet1",details);
		//WorkbookServiceImpl.protctedWorkbook("voters");
		//WorkbookServiceImpl.formatSheet("college","marksheet",semester);
		//WorkbookServiceImpl.readColumn("Student", "marksheet", 0);
		 
	}

}
