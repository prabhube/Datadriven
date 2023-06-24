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

public class WriteinExcel {

	public static void main(String[] args) throws IOException {
		File f =new File("C:\\Users\\prem\\eclipse-workspace\\DataDriven\\excel\\Book1.xlsx");
		
		FileInputStream stream=new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(stream);
		Sheet sheetAt = wb.getSheetAt(0);
		Row row = sheetAt.getRow(0);
		Cell cell = row.getCell(0);
		CellType cellTy = cell.getCellType();
		if(cellTy.equals(CellType.STRING))
		{
			String stringCellValue = cell.getStringCellValue();
			System.out.println(stringCellValue);
		}
		else if(cellTy.equals(CellType.NUMERIC))
		{
			double numericCellValue = cell.getNumericCellValue();
		int	value=(int) numericCellValue;
		System.out.println(value);
		}
		
	}
}
