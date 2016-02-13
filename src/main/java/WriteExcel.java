import java.io.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.*;
import java.util.*;


abstract class WriteExcel{
	XSSFWorkbook workbook = new XSSFWorkbook();
	XSSFSheet sheet=null;
	XSSFRow row=null;
	XSSFCell cell;
	Employee e=null;
	FileInputStream fsIP=null;

	public void writeData(Map<String,Employee> map)throws Exception{
		int j;
		System.out.println("Write working");
		fsIP= new FileInputStream(new File("EmpDetails.xlsx")); //Read the spreadsheet that needs to be updated
		XSSFWorkbook wb = new XSSFWorkbook(fsIP); 
		sheet = wb.getSheetAt(0);
		for(int i=0;i<map.size();i++){
			e=(Employee)map.get(i);
			j=0;
			row = sheet.getRow(i);
			System.out.println("Row:"+row);

			cell = row.getCell(j);
			j=j+1;
			System.out.println("Cell:"+cell);
			cell.setCellValue(e.getEmpId());
			cell=row.getCell(j);
			j=j+1;
			cell.setCellValue(e.getName());
			cell=row.getCell(j);
			j=j+1;
			cell.setCellValue(e.getSalary());
		}
		System.out.println("Data Updated to Excel");
   	        fsIP.close();
	}

	public abstract void getAllData()throws Exception;
	public abstract void getSpecificData(String userId)throws Exception;
}
