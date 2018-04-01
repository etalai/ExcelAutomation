package excelReadWrite;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelRead {
	public static void main(String[] args) throws IOException {
//		Excel hierarchy:
//			Excel Application
//				WorkBook
//					Worksheet
//						Rows
//							Cells
		String excelPath="\\Users\\etala\\Desktop\\EmpData.xlsx";
//		open the file to read
		FileInputStream in=new FileInputStream(excelPath);
//		let the Apache XSSFWorkbook class handle the data
		XSSFWorkbook workbook=new XSSFWorkbook(in);
//		jump to WorkSheet level
		XSSFSheet worksheet=workbook.getSheet("Sheet1");
		
//		find out how many rows
		int rowsCount=worksheet.getPhysicalNumberOfRows();
		System.out.println("number of rows: "+rowsCount);
//		print first row and first cell data
		System.out.println("1st row & 1st cell data: "
				+worksheet.getRow(0).getCell(0));
//		print second row and first cell data
		System.out.println("2nd row & 1st cell data: "
				+worksheet.getRow(1).getCell(0));
//		print last name using getLastRowNum() method
		System.out.println("printing the last name: "+
				worksheet.getRow(worksheet.getLastRowNum()).getCell(1));
		
		String cellValue=worksheet.getRow(worksheet.getLastRowNum()).
													getCell(1).toString();
		System.out.println(cellValue);
		System.out.println("====================");
//		print all names
		int sheetRowsCount=worksheet.getPhysicalNumberOfRows();
		
		for(int row=1; row<sheetRowsCount; row++) {
//			getting the names & assigning them to the string
			String name=worksheet.getRow(row).getCell(1).toString();
//	###		getting the names & assigning them to the string
//			using .getStringCellValue() method. this methods works if 
//			cell values are String. if they r int or char, the above 
//			method won't work
			String dept=worksheet.getRow(row).getCell(2).getStringCellValue();
//			John ---> IT
			String empID=worksheet.getRow(row).getCell(0).toString();
			System.out.println(empID+"--->"+name+"--->"+dept);	
		}
		in.close();
	}
}
