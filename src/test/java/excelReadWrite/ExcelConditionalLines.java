package excelReadWrite;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelConditionalLines {
	public static void main(String[] args) throws IOException {
		String excelPath="\\Users\\etala\\Desktop\\EmpData.xlsx";
		//open the file to read
		FileInputStream in=new FileInputStream(excelPath);
		//let the Apache XSSFWorkbook class handle the data
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		//jump to WorkSheet level
		XSSFSheet worksheet=workbook.getSheet("TestData");
		int rowsCount=worksheet.getPhysicalNumberOfRows();
//		look for searchItem that contains Cucumber
//		then mark the status as Pass
//		use for loop, with condition
		for(int rowNum=1; rowNum<rowsCount; rowNum++) {
			String str=worksheet.getRow(rowNum).getCell(1).toString();
			if(str.contains("Cucumber")) {
				XSSFCell status=worksheet.getRow(rowNum).getCell(2);
				if(status==null) {
					status=worksheet.getRow(rowNum).getCell(2);
				}
				status.setCellValue("Passed");
				break;
			}	
		}
		
		
		
		
		FileOutputStream out=new FileOutputStream(excelPath);
		workbook.write(out);
		workbook.close();	
		in.close();
	}
}
