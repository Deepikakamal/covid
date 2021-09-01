package com.excel.UserDetails;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class User_Details {
	public static void Particular_data() throws Throwable {
		File f = new File("C:\\Users\\deepi\\UserDetails\\UserDetails.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);
		Sheet sheet1 = wb.getSheetAt(0);
		Row row1 = sheet1.getRow(0);
		Cell cell1 = row1.getCell(1);
		CellType cellType1 = cell1.getCellType();
		if (cellType1.equals(CellType.STRING)) {
			String stringCellValue = cell1.getStringCellValue();
			System.out.println(stringCellValue);
			}
	else if(cellType1.equals(CellType.NUMERIC))
	{
		double numericCellValue = cell1.getNumericCellValue();
		int value = (int) numericCellValue;
		System.out.println(value);
	}
	}
	public static void main(String []args) throws Throwable
	{
		Particular_data();	
	}
	
}
