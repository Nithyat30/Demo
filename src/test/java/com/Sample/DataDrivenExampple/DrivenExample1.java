package com.Sample.DataDrivenExampple;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DrivenExample1 {

	
	public static void main(String[] args) throws IOException {
		
		File f= new File("D:\\Nithya\\Eclipse workspace\\DataDrivenExampple\\ExcelFile\\Page1.xlsx");
		FileInputStream file=new FileInputStream(f);
		Workbook wb= new XSSFWorkbook(file);
		org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheet("Sheet1");
		
		Row r=sheet.getRow(5);
		Cell cell = r.getCell(3);
		CellType cType = cell.getCellType();
		
		if(cType.equals(CellType.STRING)) {
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		}
		else if(cType.equals(CellType.NUMERIC)) {
			double numericCellValue = cell.getNumericCellValue();
			
			int value=(int)numericCellValue;
			System.out.println(value);
				
			}
		}
		
		
	}

