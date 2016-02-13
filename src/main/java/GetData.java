import java.io.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.*;
import java.util.*;


class GetData extends WriteExcel{

	FileInputStream fip=null;
	XSSFWorkbook workbook=null;
	XSSFSheet sheet=null;


	public void getAllData()throws Exception{
		fip = new FileInputStream(new File("EmpDetails.xlsx"));
		workbook = new XSSFWorkbook(fip);
		sheet = workbook.getSheetAt(0);
		for(int i = 1;i <= sheet.getLastRowNum();i++){
			row = sheet.getRow(i);
			for(int j = 0;j <= row.getLastCellNum();j++){
				cell = row.getCell(j);
				try{
					System.out.print(cell.getStringCellValue() + "\t");
				}
				catch(IllegalStateException e){
					System.out.print(cell.getStringCellValue() + "\t");
				}
			}
			System.out.println( );
		}
	}

	public void getSpecificData(String empId)throws Exception{
	
		fip = new FileInputStream(new File("EmpDetails.xlsx"));
		workbook = new XSSFWorkbook(fip);
		sheet = workbook.getSheetAt(0);
		String id=null;
		for(int i = 0;i <= sheet.getLastRowNum();i++){
			Row row = sheet.getRow(i);
			for(int j = 0;j <= row.getLastCellNum();j++){
				Cell cell = row.getCell(j);
				if(j == 0){
					id = cell.getStringCellValue();
				}
				if(id.equals(empId)){
					try{
						System.out.print(cell.getStringCellValue() + "\t");
					}
					catch(IllegalStateException e){
						System.out.print(cell.getStringCellValue() + "\t");
					}
				}
			}
    		}

	}
}




