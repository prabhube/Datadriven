package org.one.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Physical_NumberOfCell {
public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\prem\\eclipse-workspace\\DataDriven\\excel\\Book1.xlsx");
	FileInputStream stream=new FileInputStream(f);
	Workbook ws= new XSSFWorkbook(stream);
	Sheet sheet = ws.getSheet("sheet1");
	int rows = sheet.getPhysicalNumberOfRows();
	for(int i=0; i<rows;i++)
	{
	Row row = sheet.getRow(i);
	int cells = row.getPhysicalNumberOfCells();
	for(int j=0;j<cells;j++) {
		Cell cell = row.getCell(j);
		CellType cType = cell.getCellType();
		if(cType.equals(CellType.STRING))
		{
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		}
		else if(cType.equals(CellType.NUMERIC))
		{
			double numericCellValue = cell.getNumericCellValue();
			int value = (int) numericCellValue;
			System.out.println(value);
		}
	}
	}
	
	
}
}
