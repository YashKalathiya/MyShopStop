//reading value of a particular cell  
import java.io.FileInputStream;  
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.*;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.*;
import java.lang.*;
import java.io.*;

public class Prototype{  
	public String ReadCellData(Workbook wb,Sheet sheet,int vRow, int vColumn){  
	
		//variable for storing the cell value
		String value=null;        
		
		//initialize Workbook null
		wb=null;
		
		try{
			//reading data from a file in the form of bytes  
			FileInputStream fileInputStream=new FileInputStream("F:\\PrototypeData.xlsx");
			
			//constructs an XSSFWorkbook object, by buffering the whole stream into the memory  
			wb=new XSSFWorkbook(fileInputStream);  
		}
		
		catch(FileNotFoundException e) {
		
			e.printStackTrace();  
		}  
		catch(IOException e1) { 
	
			e1.printStackTrace();  
		}  
		sheet=wb.getSheetAt(0);   //getting the XSSFSheet object at given index  
		Row row=sheet.getRow(vRow); //returns the logical row  
		Cell cell=row.getCell(vColumn); //getting the cell representing the given column  
		value=cell.getStringCellValue();    //getting cell value  
		return value;               //returns the cell value  
	}
	
	public double ReadPreviousPurchase(Workbook wb,Sheet sheet,int vRow, int vColumn){  
		
		//variable for storing the cell value
		double value;        
		
		//initialize Workbook null
		wb=null;
		
		try{
			//reading data from a file in the form of bytes  
			FileInputStream fileInputStream=new FileInputStream("F:\\PrototypeData.xlsx");
			
			//constructs an XSSFWorkbook object, by buffering the whole stream into the memory  
			wb=new XSSFWorkbook(fileInputStream);  
		}
		
		catch(FileNotFoundException e) {
		
			e.printStackTrace();  
		}  
		catch(IOException e1) { 
	
			e1.printStackTrace();  
		}  
		sheet=wb.getSheetAt(0);   //getting the XSSFSheet object at given index  
		Row row=sheet.getRow(vRow); //returns the logical row  
		Cell cell=row.getCell(vColumn); //getting the cell representing the given column  
		value=cell.getNumericCellValue();   //getting cell value  
		return value;               //returns the cell value  
	}
	
	
	public static int numberOfRows(Workbook wb,Sheet sheet) {
		int no;
		wb=null; 
		String value=null; 
		try  
		{  
			//reading data from a file in the form of bytes  
			FileInputStream fileInputStream=new FileInputStream("F:\\PrototypeData.xlsx");
			
			//constructs an XSSFWorkbook object, by buffering the whole stream into the memory  
			wb=new XSSFWorkbook(fileInputStream);  
		}  
		catch(FileNotFoundException e){  
		
			e.printStackTrace();  
		}  
		catch(IOException e1) { 
		  
			e1.printStackTrace();  
		}  
		
		sheet=wb.getSheetAt(0);
		no = sheet.getPhysicalNumberOfRows();
		return no;
	}
	public static void main(String[] args) throws IOException{
		try {
		String userName,userPassword;
		String dataBaseUser="",dataBasePass="";
		int numberOfRow,flag = 0;
		String password;
		Workbook wb = null;
		
		Row row;
		Cell cell;
		Scanner sc = new Scanner(System.in);
		
		//System.out.println("Number of records in Sheet :"+(numberOfRow-1));
		FileInputStream fileInputStream=new FileInputStream("F:\\PrototypeData.xlsx");
		//Writing data to a file in the form of bytes
		wb = new XSSFWorkbook(fileInputStream);
		Sheet sheet=wb.getSheetAt(0);
		numberOfRow = numberOfRows(wb,sheet);
		
		System.out.println("Enter User name:");
		userName = sc.nextLine();
		System.out.println("Enter Password:");
		password = sc.nextLine();
		
		Prototype rc=new Prototype();   //object of the class 
	
		//reading the value of i^th row and 1st column
		for(int i=1 ; i<numberOfRow ; i++) {
			dataBaseUser=rc.ReadCellData(wb,sheet,i,0);
			dataBasePass=rc.ReadCellData(wb,sheet,i,1);
			if(dataBaseUser.equals(userName) && dataBasePass.equals(password)) {
				//System.out.println("User "+dataBaseUser+" Found!!!");
				//Getting previous purchase records if user has any.
				double previousPurchase = rc.ReadPreviousPurchase(wb,sheet,i,2);
				//System.out.println(previousPurchase);
				if(previousPurchase>1500.0) {
					System.out.println("Welcome "+dataBaseUser+"!");
					System.out.println("We value your membership with us");
					System.out.println("Congratulations.... You have won Giftcard worth 50$.");
				}
				else {
					System.out.println("Keep shopping and get exciting offers every month!!!!");
				}
				break;
			}
			else {
				flag = 1;
			}
		}
		
		if(flag == 1) {
			System.out.println("Looks Like you are not a registered member,\n Register now and get 40% discount on first 3 purchases.....\n 1 - Yes \n 2 - No");
			
			int op;
			op = sc.nextInt();
			if(op==1) {
				
				FileOutputStream fileOutputStream = new FileOutputStream(new File("F:\\PrototypeData_1.xlsx"));
				wb = new XSSFWorkbook();
				wb.write(fileOutputStream);
				
				row     = sheet.createRow(numberOfRow);  
		        cell   = row.createCell(0);  
		        //int lastRow = sheet.getPhysicalNumberOfRows();
		        sc.nextLine();
		        System.out.println("Enter User name:");
				userName = sc.nextLine();
		        cell.setCellValue(userName);
		        
		        System.out.println("Enter your password: ");
		        password = sc.nextLine();
		        cell = row.createCell(1);
		        cell.setCellValue(password);
		        
		        cell = row.createCell(1);
		        cell.setCellValue(0);
		        System.out.println("Registered Successfully!!!");
		        fileOutputStream.close();       
			}
			
			if(op==2) {
				System.out.println("Register with us and find exciting deals every month. Thank you!");
				System.exit(0);
			}
			
		}
		
		}
		catch(Exception e) {
			System.out.println(e);
		}

	}
}