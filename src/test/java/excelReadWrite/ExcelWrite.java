package excelReadWrite;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public static void main(String[] args) throws IOException {
		String excelPath="\\Users\\etala\\Desktop\\EmpData.xlsx";
		//open the file to read
		FileInputStream in=new FileInputStream(excelPath);
		//let the Apache XSSFWorkbook class handle the data
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		//jump to WorkSheet level
		XSSFSheet worksheet=workbook.getSheet("Sheet1");
		int rowsCount=worksheet.getPhysicalNumberOfRows();
		System.out.println(rowsCount);
		
		XSSFCell cell=worksheet.getRow(1).getCell(2);
		if(cell==null) {
			cell=worksheet.getRow(1).createCell(2);
		}
		cell.setCellValue("Pass");
//		Fail the android
//		cell=worksheet.getRow(5).getCell(2);
//		if(cell==null) {
//			cell=worksheet.getRow(5).getCell(2);
//		}
//		cell.setCellValue("Fail");
		
		
//	###	below lines of "out" are always stand at the end.
//	###	Close excelSheet before running the class
//		workbook.close();
		FileOutputStream out=new FileOutputStream(excelPath);
		in.close();
		out.close();	
		
	}
}
