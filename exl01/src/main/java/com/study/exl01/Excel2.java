package com.study.exl01;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import javax.swing.JFileChooser;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class Excel2 {
	
	@Test   //读取表格里的内容
	public void test1() throws FileNotFoundException, IOException{
		JFileChooser fileChooser = new JFileChooser();
		
		int returnValue = fileChooser.showOpenDialog(null);
		
		
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			
			Workbook workbook = new HSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
			Sheet sheet = workbook.getSheetAt(0);
			
			for(Iterator<Row> rit=sheet.rowIterator();rit.hasNext();){
				Row row = rit.next();
				
				for(Iterator<Cell> cit=row.cellIterator();cit.hasNext();){
					Cell cell = cit.next();
					cell.setCellType(Cell.CELL_TYPE_STRING);
					System.out.print(cell.getStringCellValue() + "\t\t");
				}
				System.out.println();
			}
			
			
			
		}
		
		
		
	}

}
