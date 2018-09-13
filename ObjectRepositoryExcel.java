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

public class ObjectRepositoryExcel {
	
	/*
public static void main(String[] args) {
	
		
	//Map<String, String> s=excelRead("C:\\Users\\rakeshkumar.t\\sahi_pro\\userdata\\scripts\\object_parm.xlsx","objectid");
	Map<String, String> s=excelRead("C:\\Users\\rakeshkumar.t\\sahi_pro\\userdata\\scripts\\IBM_RO_JY_7.2_SP3_Eng_HCL\\IBM_RO_HCL_Automation\\Conf\\Vault_Creation_Enhancement_Collections\\vault_creation_ObjectRepo.xlsx","Vault_Config");
		System.out.println(s.get("ObjDiscover"));
		
	}
	*/

	
	public static Map<String, String> excelRead(String FilePath , String Sheetname){
		Map<String, String> columnMap = new LinkedHashMap<String, String>();
try {
	
	FileInputStream file=new FileInputStream(FilePath);
	//file.setWritable(false);
Workbook workbook = WorkbookFactory.create(file);
        
			Sheet sheet = workbook.getSheet(Sheetname);
             for (int k=1;k<=sheet.getLastRowNum();k++){
				Row r = sheet.getRow(k);
				Iterator<Cell> i = r.cellIterator();				
				Cell i2 =r.cellIterator().next();
				Cell c=null;
				
			      while(((i.hasNext()))) {
					 c = i.next();
					//Cell  c1=i2.next();				
					
					}
			      String c1=i2.getStringCellValue();
			      columnMap.put(c1.toString().trim(),c.getStringCellValue().trim());
				
			}
             
             System.out.println(">>>>>>>>>>>"+ columnMap);
             
			} 
 
catch (Exception e) {
			e.printStackTrace();
		}
return columnMap;
 }
		
	}


