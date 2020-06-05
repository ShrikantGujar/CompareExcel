package com.pwc.automation.comapre_sheet;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import com.relevantcodes.extentreports.NetworkMode;


public class CompareExcel {
	
	public void comapreExcel() throws IOException {
		
		ExtentReports reports = new ExtentReports("Report/report.html", true, NetworkMode.OFFLINE);
		ExtentTest extentTes = new ExtentTest("Report Difference", "Please refer below - ");
		
		
		FileInputStream inputStream1 = new FileInputStream("");
		HSSFWorkbook book1 = new HSSFWorkbook(inputStream1);
		HSSFSheet book1Sheet1 = book1.getSheetAt(0);
		int row1 = book1Sheet1.getPhysicalNumberOfRows();
		
		FileInputStream inputStream2 = new FileInputStream("");
		HSSFWorkbook book2 = new HSSFWorkbook(inputStream2);
		HSSFSheet book2Sheet1 = book1.getSheetAt(0);
		int row2 = book1Sheet1.getPhysicalNumberOfRows();
		
		if(row1 == row2) {
			
			for(int index = 0; index < row1 ; index++) {
				HSSFRow row11 = book1Sheet1.getRow(index);
				HSSFRow row22 = book2Sheet1.getRow(index);
				
				String idstr1 = "";
				
				HSSFCell id1 = row11.getCell(0);
				
				if(id1 != null) {
					id1.setCellType(CellType.STRING);
					idstr1 = id1.getStringCellValue();
				}
				
				
				String idstr2 = "";
				
				HSSFCell id2 = row11.getCell(0);
				
				if(id2 != null) {
					id2.setCellType(CellType.STRING);
					idstr2 = id2.getStringCellValue();
				}
				
				if(!idstr1.equals(idstr2)) {
					System.out.println("not matched");
					extentTes.log(LogStatus.ERROR, idstr1 + " not mateched " + idstr2);
				}
			}
		}
	}
}
