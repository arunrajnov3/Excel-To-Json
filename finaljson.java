package com.howtodoinjava.demo.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;


public class finaljson
{
	public static void main(String[] args) 
	{
		int ctr=0,c=1;
		int i=1;
		try
		{
			FileInputStream file = new FileInputStream(new File("pro.xlsx"));
			int maxrow=rowmax();
			int maxcol=colmax();
			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			
//			map.put("Name", "www.javainterviewpoint.com");
			FileWriter fileWriter = new FileWriter("f:\\sample.json");
//			JSONObject jsonObject= new JSONObject();
			
			while (rowIterator.hasNext()) 
			{
				ctr=0;
				JSONObject jsonObject= new JSONObject();
				Row row = rowIterator.next();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				JSONArray jsonArray = new JSONArray();
				while (cellIterator.hasNext()) 
				{
					char a='A';
					Cell cell = cellIterator.next();
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
						case Cell.CELL_TYPE_NUMERIC:
							System.out.print(cell.getNumericCellValue() + "\t");
							//jsonArray.put(cell.getNumericCellValue() + "\t");
							break;
						case Cell.CELL_TYPE_STRING:
							System.out.print(cell.getStringCellValue() + "\t");
							jsonArray.put( cell.getStringCellValue());
							
							if(ctr==0)
							{
								jsonObject.put( "A",cell.getStringCellValue());
							}
							if(ctr==1)
							{
								jsonObject.put( "B",cell.getStringCellValue());
							}
							if(ctr==2)
							{
								jsonObject.put( "C",cell.getStringCellValue());
							}if(ctr==3)
							{
								jsonObject.put( "D",cell.getStringCellValue());
							}if(ctr==4)
							{
								jsonObject.put( "E",cell.getStringCellValue());
							}if(ctr==5)
							{
								jsonObject.put( "F",cell.getStringCellValue());
							}
						
							break;
					}
					ctr++;
				}
				//jsonObject.put("A", jsonArray);
				if(c==1){
				fileWriter.write("["+jsonObject.toString()+",");
				}
				else if(c==maxcol)
				{
					fileWriter.write(jsonObject.toString()+"]");
				}
				else
				{
					fileWriter.write(jsonObject.toString()+",");
				}
	            
				System.out.println("");
				c++;
				
			}
			fileWriter.close();

            System.out.println("JSON Object Successfully written to the file!!");
//			FileWriter fileWriter = new FileWriter("f:\\sample.json");
//
//            // Writting the jsonObject into sample.json
//            fileWriter.write(jsonObject.toString());
//            fileWriter.close();
//
//            System.out.println("JSON Object Successfully written to the file!!");
			file.close();
			
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}//main
	public static int rowmax()
	{
		int rows=0;
		try
		{
			FileInputStream file = new FileInputStream(new File("Book1.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) 
			{
				Row row = rowIterator.next();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
						case Cell.CELL_TYPE_NUMERIC:
							break;
						case Cell.CELL_TYPE_STRING:
							rows++;
							break;
					}
				}
				break;
			}
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		return rows;
	}//row max
	
	
	public static int colmax()
	{
		int cols=0;
		try
		{
			FileInputStream file = new FileInputStream(new File("Book1.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			//Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) 
			{
				Row row = rowIterator.next();
				//For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
						case Cell.CELL_TYPE_NUMERIC:
							break;
						case Cell.CELL_TYPE_STRING:
							break;
					}
				}
				cols++;
			}
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		return cols;
	}//row max
	
	
}
