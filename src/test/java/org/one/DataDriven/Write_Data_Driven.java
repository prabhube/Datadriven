package org.one.DataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Data_Driven {
	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\prem\\eclipse-workspace\\DataDriven\\excel\\Book1.xlsx");
		FileInputStream stream=new FileInputStream(f);
		Workbook wb=new XSSFWorkbook(stream);
		//wb.createSheet("data1").createRow(0).createCell(0).setCellValue("email");
		wb.getSheet("data1").getRow(0).createCell(1).setCellValue("Password");
		wb.getSheet("data1").createRow(1).createCell(0).setCellValue("value");
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		System.out.println("hai");
	}
}
