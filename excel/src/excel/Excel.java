package excel;

import java.io.FileInputStream;

import jxl.Sheet;
import jxl.Workbook;
public class Excel {

	public static void main(String[] args) throws Exception{
		FileInputStream f1=new FileInputStream("E:\\Selenium\\ReadExcel.xls");
		Workbook w1=Workbook.getWorkbook(f1);
		Sheet s1=w1.getSheet("Sheet");
		System.out.println(s1.getName());
		int i=2;
		String EmpId=s1.getCell(0,i).getContents();
		String EmpName=s1.getCell(1,i).getContents();
		String EmpSal=s1.getCell(2,i).getContents();
		System.out.println(EmpId+"\n"+EmpName+"\n"+EmpSal);
	}
}
