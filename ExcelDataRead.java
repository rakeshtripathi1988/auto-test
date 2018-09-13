package com.test;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ddf.EscherColorRef.SysIndexSource;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataRead {

/*
public static void main(String[] args) {
	
		//ExcelDataRead ob=new ExcelDataRead();
		List s=excelRead("C:\\Users\\rakeshkumar.t\\sahi_pro\\userdata\\scripts\\Vault_Parm.xlsx","APP_PFR_Sol_Parameter");
		//System.out.println(">>>" + s.get(0).get("vaulttype"));
		Map<String, String> columnMap1=(Map<String, String>) s.get(1);
		System.out.println(columnMap1.get("CompIPPR"));
		
	}
	*/

	 
	public static List excelRead(String FilePath , String Sheetname){
		List<Map<String, String>> adddata=new ArrayList<Map<String, String>>();
try {
	
	FileInputStream file=new FileInputStream(FilePath);
Workbook workbook = WorkbookFactory.create(file);
			Sheet sheet = workbook.getSheet(Sheetname);
             for (int k=1;k<=sheet.getLastRowNum();k++){
				Row r1 = sheet.getRow(0);
				Iterator<Cell> i1 = r1.cellIterator();
				Row r = sheet.getRow(k);
				Iterator<Cell> i = r.cellIterator();
					//Cell c1=null;
				Map<String, String> columnMap = new LinkedHashMap<String, String>();
			      while(i.hasNext()) {
					Cell c = i.next();
					Cell  c1=i1.next();
					columnMap.put(c1.getStringCellValue().trim(), c.getStringCellValue().trim());
					}
				//System.out.println("\n");	
				
				//System.out.println("" + columnMap.keySet());
				adddata.add(columnMap);	
				
			}
             
             
			} 
 
catch (Exception e) {
			e.printStackTrace();
		}
return adddata;
 }
		
	}


