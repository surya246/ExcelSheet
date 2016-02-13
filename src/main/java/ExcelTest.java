import java.io.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.*;
import java.util.*;



public class ExcelTest{

	public static void main(String[] args)throws Exception{
		HashMap<String,Employee> map=new HashMap<String,Employee>();
		WriteExcel d=new GetData();
	        //System.out.println(d.toString());
		Scanner s =new Scanner(System.in);
		String choice;
		System.out.print("For how many employees you want to store data:");
		int count=s.nextInt();
		String empid;
		int j=0;
		Employee e=null;
		for(int i=1;i<=count;i++){
		e=new Employee();
			System.out.print("Enter "+i+":Employee ID:");
			empid=s.next();
			e.setEmpId(empid);
			System.out.print("Enter "+i+":Employee Name:");
			e.setName(s.next());
			System.out.print("Enter "+i+":Employee Salary:");
			e.setSalary(s.next());
			map.put(empid,e);
		}

		d.writeData(map);
		System.out.println("Do you want to read data(YES/NO)/:");
		choice=s.next();
		if(choice.equals("YES")){
			System.out.print("You want to read total data (YES/NO)");
			choice = s.next();
			if(choice.equals("YES")){
				d.getAllData();
			}
			else if(choice.equals("NO")){
				System.out.print("Enter Id:");
				d.getSpecificData(s.next());
			}
			else{
				System.out.println("Invalid input enter YES/NO");
			}
		}

	}
}

