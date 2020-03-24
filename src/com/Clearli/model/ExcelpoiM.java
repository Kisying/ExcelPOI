package com.Clearli.model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;

import org.apache.poi.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.openxml4j.opc.*;

public class ExcelpoiM {
	private final static String xls = "xls";  
	 private final static String xlsx = "xlsx";  
	// �C�L Excel �\�� �p�p�p�p�p�p�p�p
	public void PutDataToColunm(Sheet sheet,String KEYS,String Value) throws IOException {
		String Cols = "",Rows = "";
		for(int i=0;i<KEYS.trim().length();i++){
			String T = KEYS.trim().substring(i,i+1);
			if(checkisNum(T.trim())) Rows += T.trim();
			else{
				Cols += T.trim().toUpperCase();
			}
		}
		int Col = -1;
		if(Cols.trim().length()==1){
			Col = (int)Cols.trim().charAt(0);
			Col = Col - 64;
		}
		else if(Cols.trim().length()==2){
			int Col1 = (int)Cols.trim().charAt(0);
			int Col2 = (int)Cols.trim().charAt(1);
			Col1 = Col1 - 64;
			Col2 = Col2 - 64;
			Col = (26*Col1) + Col2;
		}
		sheet.getRow(Integer.parseInt(Rows)-1).getCell(Col-1).setCellValue(Value.trim());
	}
	public static Workbook wbcreate(File file) throws Exception {
		
		String fileName = file.getName();
		InputStream in = new FileInputStream(file);
		
		if (!in.markSupported()) {
			in = new PushbackInputStream(in, 8);
		}
		if(fileName.endsWith(xls)){
			return new HSSFWorkbook(in);
		}
		if(fileName.endsWith(xlsx)){
			return new XSSFWorkbook(in);
			//return new XSSFWorkbook(OPCPackage.open(in));
		}
		return null;
	}
	public static byte[] getBytesFromFile(File file) throws IOException {
		InputStream is = new FileInputStream(file);
		long length = file.length();
		byte[] bytes = new byte[(int)length];
		int offset = 0;
		int numRead = 0;
		while (offset < bytes.length && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0){
			offset += numRead;
		}
		is.close();
		return bytes;
	}
	//�ˮּƦr�榡
	public boolean checkisNum (String num) throws IOException {
		try  
		  {  
		    double d = Double.parseDouble(num);  
		  }  
		  catch(NumberFormatException nfe)  
		  {  
		    return false;  
		  }  
		  return true;  
	}
	//�P�_���O�_�s�b  
	public static void checkFile(File file) throws IOException{  
	 
		if(null == file){  
			System.out.println("��󤣦s�b�I");  
			throw new FileNotFoundException("��󤣦s�b�I");  
		}  
		//��o���W  
		String fileName = file.getName(); 
		//�P�_���O�_�Oexcel���  
		if(!fileName.endsWith(xls) && !fileName.endsWith(xlsx)){  
			System.out.println(fileName + "���Oexcel���");  
			throw new IOException(fileName + "���Oexcel���");  
		}  
	}
	
}
