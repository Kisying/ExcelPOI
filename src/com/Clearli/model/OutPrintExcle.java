package com.Clearli.model;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.openxml4j.opc.*;
import com.Clearli.model.*;
import org.apache.poi.hssf.record.*;

public class OutPrintExcle {
	
	static ExcelpoiM ExcelMod = new ExcelpoiM();
	
	public static void main (String InAdress,String OutArdess) throws Exception{
		
		 ArrayList <String> nameList = new ArrayList(); 
		
		for (File f : new File(InAdress).listFiles()) {
			String f_s = f.toString();
			
			String Fils = f_s.substring(f_s.lastIndexOf("\\")+1,f_s.lastIndexOf("xls")-1);
			nameList.add(Fils);
			// System.out.println(Fils);
		}
		for (int i=0; i<nameList.size(); i++) {//nameList.size()
			
            //System.out.println(nameList.get(i));
            
			String file = nameList.get(i);
			String ExcelFileName= file.substring(0,file.indexOf(".")+1);
			
			File files = new File (InAdress+"\\"+file+".xls");
			String absName = files.getAbsolutePath();
			//System.out.println("absName==>"+absName.trim());
			FileInputStream fis = new FileInputStream (files);
			Workbook workbook1 = ExcelMod.wbcreate(files);
			Workbook workbook2 = ExcelMod.wbcreate(files);
			
			workbook1.removeSheetAt(workbook1.getSheetIndex("下行")); 
			workbook2.removeSheetAt(workbook2.getSheetIndex("上行")); 
			
			changeExcel(workbook1);
			changeExcel(workbook2);
			newfild(OutArdess,ExcelFileName);
			
			String Ex_U = OutArdess+"/"+ExcelFileName+"/"+file+".U.xls" ;
			String Ex_D = OutArdess+"/"+ExcelFileName+"/"+file+".D.xls";
			
			
			FileOutputStream out1 = new FileOutputStream(new File(Ex_U));
			workbook1.write(out1);
			out1.close();
			
			FileOutputStream out2 = new FileOutputStream(new File(Ex_D));
			workbook2.write(out2);
			out2.close();
		}
	}

	public static void changeExcel(Workbook workbook) throws IOException {
		
		Sheet sheet = workbook.getSheetAt(0);
		Font fontBlack = workbook.createFont();
		fontBlack.setFontName("Arial");
		fontBlack.setFontHeightInPoints((short)10);
		fontBlack.setColor(HSSFColor.BLACK.index);
		Cell cell;
		
		String ary[] = new String [7];
		Row  row = sheet.getRow(0);
		if(row!=null){
			 for (int j = 2; j < 8; j++){  //看資料需要的欄數 
				 cell = row.getCell(j); 
				 
				 if(cell == null)
			      {
					// System.out.println("Null");
			      }else {
			    	  ary[j-2] = cell.toString();
			    	// System.out.println(cell.toString());//取出j列j行的值                
			      }
				
			 } 
		 } 
	
			ExcelMod.PutDataToColunm(sheet,"A1",ary[0]);
			ExcelMod.PutDataToColunm(sheet,"B1",ary[1]);
			ExcelMod.PutDataToColunm(sheet,"C1",ary[2]);
			ExcelMod.PutDataToColunm(sheet,"D1",ary[3]);
			ExcelMod.PutDataToColunm(sheet,"E1",ary[4]);
			ExcelMod.PutDataToColunm(sheet,"F1",ary[5]);
			ExcelMod.PutDataToColunm(sheet,"G1","");
			ExcelMod.PutDataToColunm(sheet,"H1","");
	}
	
	public static void newfild (String address,String file) {
		File file_add= new File(address+"\\"+file);

		/* 檔案 */

		if (!file_add.exists()) { // 不存在檔案時，建立
			file_add.mkdir();

		} else { // 檔案存在時刪除

			file_add.delete();
			file_add.mkdir();
		}
	}
	
}
