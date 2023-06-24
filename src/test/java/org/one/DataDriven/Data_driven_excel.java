package org.one.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data_driven_excel {
	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\prem\\eclipse-workspace\\DataDriven\\excel\\Book2.xlsx");
		FileInputStream stream= new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(stream);
		wb.getSheet("sheet2").createRow(0).createCell(0).setCellValue("name");
		wb.getSheet("sheet2").getRow(0).createCell(1).setCellValue("company");
		wb.getSheet("sheet2").createRow(1).createCell(0).setCellValue("prabhu");
		wb.getSheet("sheet2").getRow(1).createCell(1).setCellValue("ispirisys");
		wb.getSheet("sheet2").createRow(2).createCell(0).setCellValue("prem");
		wb.getSheet("sheet2").getRow(2).createCell(1).setCellValue("itc");
		
		FileOutputStream fos=new FileOutputStream(f);
		wb.write(fos);
		
		Sheet sheet = wb.getSheet("sheet2");
		int NumberOfRows = sheet.getPhysicalNumberOfRows();
		for(int i=0;i<=NumberOfRows;i++)
		{
			Row row = sheet.getRow(i);
			int NumberOfCells = row.getPhysicalNumberOfCells();
			for(int j=0;j<=NumberOfCells;j++)
			{
				Cell cell = row.getCell(j);
				CellType ceType = cell.getCellType();
				if(ceType.equals(CellType.STRING))
				{
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				}
				else if(ceType.equals(CellType.NUMERIC))
				{
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;
					System.out.println(value);
				}
					
			}
		}
		
	}
	
			
}
