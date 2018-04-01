package excelReadWrite;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.ListDocument.List;
public class ExcelConditionalRead {

	public static void main(String[] args) throws IOException {
		String excelPath="\\Users\\etala\\Desktop\\EmpData.xlsx";
		//open the file to read
		FileInputStream in=new FileInputStream(excelPath);
		//let the Apache XSSFWorkbook class handle the data
		XSSFWorkbook workbook=new XSSFWorkbook(in);
		//jump to WorkSheet level
		XSSFSheet worksheet=workbook.getSheet("TestData");
		int rowsCount=worksheet.getPhysicalNumberOfRows();
		
		for(int row=1; row<rowsCount; row++) {
			String execute=worksheet.getRow(row).getCell(0).toString();
//		read first cell value
//		if it is YES, switch to the cell and print the reach item
			if(execute.equals("Yes")){
				String searchItem=worksheet.getRow(row).getCell(1).toString();
				System.out.println("Searching for: "+searchItem);
			}
		}	
		in.close();
	}
}
