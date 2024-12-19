package excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WriteExcelData {

	public static void main(String[] args) {
		FileInputStream fis=new FileInputStream("D:\\Selenium\\WriteData.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheet("test");
		System.out.println(sheet.getSheetName());
		System.out.println(sheet.getLastRowNum());
		System.out.println("Before updating Cell Data"+sheet.getRow(2).getCell(1));
		XSSFCell cell=sheet.getRow(2).getCell(1);
		cell.setCellValue("Test123");
		fis.close();
		FileOutputStream fileOut=new FileOutputStream("D:\\Selenium\\WriteData.xlsx");
		workbook.write(fileOut);
		System.out.println("Updated data after write is done "+cell.getStringCellValue());
		fileOut.close();
	}
}
