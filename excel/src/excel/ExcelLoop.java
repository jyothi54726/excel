package excel;

import java.io.FileInputStream;

import jxl.Sheet;
import jxl.Workbook;

public class ExcelLoop {

	public static void main(String[] args) throws Exception {
		FileInputStream f1=new FileInputStream("E:\\Selenium\\ReadExcel.xls");
		Workbook objwb=Workbook.getWorkbook(f1);
		Sheet s1=objwb.getSheet(0);
		int rows=s1.getRows();
		System.out.println(rows);
		for(int i=1;i<rows;i++) {
			String EmpID=s1.getCell(0,i).getContents();
			String EmpName=s1.getCell(1,i).getContents();
			String EmpSal=s1.getCell(2,i).getContents();
			System.out.println(EmpID+"\n"+EmpName+"\n"+EmpSal);
		}
	}
}
