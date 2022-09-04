/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */

package com.mycompany.readxlsxfile;
// Java Program to Illustrate Reading
// Data to Excel File Using Apache POI

// Import statements

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Main class
public class ReadXLSXfile {
        
	// Main driver method
	public static void main(String[] args)
	{       
                
                System.out.println("Enter the number of students");
                Scanner myObj = new Scanner(System.in);
                int i = myObj.nextInt()+3;
                System.out.println("Enter number of subjects");
                Scanner myObj2 = new Scanner(System.in);
                int j = myObj2.nextInt();
                int array[][] = new int[i+2][j+2];
                int total[] = new int[i];
		// Try block to check for exceptions
		try {

			// Reading file from local directory
			FileInputStream file = new FileInputStream(
				new File("C:\\Users\\ANUJ\\Downloads\\JavaTest.xlsx"));

			// Create Workbook instance holding reference to
			// .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
                        int p =0;
			// Till there is an element condition holds true
			while (rowIterator.hasNext()) {
                                
				Row row = rowIterator.next();

				// For each row, iterate through all the
				// columns
				Iterator<Cell> cellIterator
					= row.cellIterator();
                                int q = 0;
                                total[p]=0;
				while (cellIterator.hasNext()) {
                                        
					Cell cell = cellIterator.next();

					// Checking the cell type and format
					// accordingly
					if(cell.getCellType()==CellType.NUMERIC)
                                        {
                                                array[p][q] = (int) cell.getNumericCellValue();
						System.out.print(
							cell.getNumericCellValue()
							+ "\t"+"    ");
                                                while(q!=0)
                                                {
                                                total[p]=total[p]+array[p][q];
                                                System.out.print("    ");
                                                break;
                                                }
                                                
                                        }

					if(cell.getCellType()==CellType.STRING)
                                        {
						System.out.print(
							cell.getStringCellValue()
							+ "\t");
					}
                                        q=q+1;  
				}
                                
                                int z=p-1;
                                
                                while(z>0)
                                {
                                System.out.println("Marks obtained by "+z+" is "+total[p]);
                                break;
                                }
                                p=p+1;
                                
                                
                                System.out.println("");
			}

			// Closing file output streams
			file.close();
		}

		// Catch block to handle exceptions
		catch (Exception e) {

			// Display the exception along with line number
			// using printStackTrace() method
			e.printStackTrace();
		}
	}
}Anuj