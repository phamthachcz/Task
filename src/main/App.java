package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;





public class App {
	
	
	public static void main(String[] args) {
		
		//Input example: D:\\Data.xlsx
		Scanner sc= new Scanner(System.in);  
		System.out.println("Enter the path of file: ");
		String pathString = sc.nextLine().trim();
		readXLSXFile(pathString);
		sc.close();
	}
	
	//Check number is Prime or not
	public static Boolean isPrime(int number) {
		if(number <= 1) {
			return false;
		}
		for(int i = 2; i < Math.sqrt(number); i++) {
			if(0 == number % i) {
				return false;
			}
		}
		return true;
	}
	
	//Process file xlsx and check numbers
	private static void readXLSXFile(String fileName) {
		// TODO Auto-generated method stub
		try {
			File inputFile = new File(fileName);
			FileInputStream fileInputStream = new FileInputStream(inputFile);
			XSSFWorkbook work = new XSSFWorkbook(fileInputStream);
			
			//List result about primes in file
			List<Integer> primes = new ArrayList<Integer>();
			//Column Data index
			int colIndex = 1;
			
			Iterator<Sheet> sheetIterator = work.sheetIterator();
			
			while(sheetIterator.hasNext()) {
				Sheet sheet = sheetIterator.next();
				for(int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
					Row row = sheet.getRow(rowIndex);
					if(row != null) {
						Cell cell = row.getCell(colIndex);
						if(cell != null) {
							String cellValue = cell.getStringCellValue();
							if(!cellValue.isEmpty()) {
								try{
						            int number = Integer.parseInt(cellValue);
						            if(number > 0) {
						            	if(isPrime(number)) {
						            		primes.add(number);
						            	}
						            }
						             
						        }
						        catch (NumberFormatException ex){
						            
						        }
							}
						}
					}
				}
			}
			//Sort order result
			Collections.sort(primes);
			
			//print result
			String primesString = primes.toString();
			primesString = primesString.substring(1, primesString.length() - 1);
			System.out.println(primesString);
			
			//Close file
			work.close();
			fileInputStream.close();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println("File input not found!");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
}
