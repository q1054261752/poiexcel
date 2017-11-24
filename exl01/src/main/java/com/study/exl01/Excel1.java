package com.study.exl01;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.junit.Test;

/**
 * Hello world!
 *
 */
public class Excel1 
{
	@Test //excell   sheet测试
    public  void test1()
    {
    	Workbook workbook = new HSSFWorkbook();
    	
    	Sheet sheet1 = workbook.createSheet();
    	Sheet Sheet2 = workbook.createSheet("第一个sheet1");
    	Sheet Sheet3 = workbook.createSheet("第二个sheet2");
    	Sheet Sheet4 = workbook.createSheet(WorkbookUtil.createSafeSheetName("[]sdf?sdfsd你好"));//用WorkbookUtil过滤特殊字符
    	try {
        	FileOutputStream output = new FileOutputStream("Test1.xls");
			workbook.write(output);
			output.close();
    	
    	} catch (Exception e) {
			e.printStackTrace();
		}
    }
	
	@Test //excell   sheet测试
    public  void test2()
    {
    	Workbook workbook = new HSSFWorkbook();
    	Sheet sheet = workbook.createSheet("Eggs");
//    	Row row = sheet.createRow(0);
//    	Cell cell = row.createCell(3);  合起来写
    	Cell cell = sheet.createRow(0).createCell(3);
    	cell.setCellValue("Hi there");
    	try {
			FileOutputStream output = new FileOutputStream("Test2.xls");
			workbook.write(output);
			output.close();
    	} catch (Exception e) {
			e.printStackTrace();
		}
    }
	
	@Test //excell   获取cell上的文本信息
    public  void test3()
    {
    	Workbook workbook = new HSSFWorkbook();
    	Sheet sheet = workbook.createSheet("Eggs");
//    	Row row = sheet.createRow(0);
//    	Cell cell = row.createCell(3);  合起来写
    	Cell cell = sheet.createRow(0).createCell(3);
    	cell.setCellValue("Hi there");
    	
    	System.out.println(cell.getRichStringCellValue().toString());//获取cell上的文本信息
    	
    	try {
			FileOutputStream output = new FileOutputStream("Test2.xls");
			workbook.write(output);
			output.close();
    	} catch (Exception e) {
			e.printStackTrace();
		}
    }
	
	@Test //excell   计算工式
    public  void test4()
    {
    	Workbook workbook = new HSSFWorkbook();
    	Sheet sheet = workbook.createSheet("Eggs");

    	Cell cell1 = sheet.createRow(0).createCell(0);
    	Cell cell2 = sheet.createRow(0).createCell(1);
    	Cell cell3 = sheet.createRow(0).createCell(2);
    	Cell cell4 = sheet.createRow(0).createCell(3);
    	Cell cell5 = sheet.createRow(0).createCell(4);
    	
    	cell1.setCellValue(56);
    	cell2.setCellValue("+");
    	cell3.setCellValue(199);
    	cell4.setCellValue("=");
//    	cell5.setCellValue("=A1+C1"); //不对
    	cell5.setCellFormula("SUM(A1:C1)");
    	
    	try {
			FileOutputStream output = new FileOutputStream("Test2.xls");
			workbook.write(output);
			output.close();
    	} catch (Exception e) {
			e.printStackTrace();
		}
    }
	

	@Test //excell   创建   excell style 
    public  void test5()
    {
    	Workbook workbook = new HSSFWorkbook();
    	Sheet sheet = workbook.createSheet("Eggs");

    	Cell cell1 = sheet.createRow(0).createCell(0);
    	
    	//设置背景的颜色
    	CellStyle style = workbook.createCellStyle();
    	style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
    	style.setFillPattern(CellStyle.FINE_DOTS);
    	
    	//设置字体
    	Font font = workbook.createFont();
    	font.setColor(IndexedColors.YELLOW.getIndex());
    	font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    	font.setItalic(true);
//    	font.setUnderline(Font.U_DOUBLE);
    	font.setFontName("Helvetia");
    	style.setFont(font);
    	
    	cell1.setCellStyle(style);
    	cell1.setCellValue("我爱我家！");
    	
    	try {
			FileOutputStream output = new FileOutputStream("Test2.xls");
			workbook.write(output);
			output.close();
    	} catch (Exception e) {
			e.printStackTrace();
		}
    }
	
	@Test //excell   设置列的宽高
    public  void test6()
    {
    	Workbook workbook = new HSSFWorkbook();
    	Sheet sheet = workbook.createSheet("Eggs");
    	sheet.setColumnWidth(0, 2000);   //设置单元格的宽

    	Cell cell = sheet.createRow(0).createCell(0);
//    	cell.getRow().setHeightInPoints(30);    //设置单元格的高
    	cell.getRow().setHeight((short)600);//1/20th of a point.
    	
    	//设置背景的颜色
    	CellStyle style = workbook.createCellStyle();
    	style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
    	style.setFillPattern(CellStyle.FINE_DOTS);
    	
    	//设置字体
    	Font font = workbook.createFont();
    	font.setColor(IndexedColors.YELLOW.getIndex());
    	font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    	font.setItalic(true);
//    	font.setUnderline(Font.U_DOUBLE);
    	font.setFontName("Helvetia");
    	style.setFont(font);
    	
    	cell.setCellStyle(style);
    	cell.setCellValue("我爱我家！");
    	
    	try {
			FileOutputStream output = new FileOutputStream("Test2.xls");
			workbook.write(output);
			output.close();
    	} catch (Exception e) {
			e.printStackTrace();
		}
    }
	
	@Test //excell   设置列的合并单元格
    public  void test7()
    {
    	Workbook workbook = new HSSFWorkbook();
    	Sheet sheet = workbook.createSheet("Eggs");
    	sheet.addMergedRegion(new CellRangeAddress(/*Row*/0,4,/*Column*/0,3));
    	sheet.setColumnWidth(0, 2000);   //设置单元格的宽

    	Cell cell = sheet.createRow(0).createCell(0);
//    	cell.getRow().setHeightInPoints(30);    //设置单元格的高
    	cell.getRow().setHeight((short)600);//1/20th of a point.
    	
    	//设置背景的颜色
    	CellStyle style = workbook.createCellStyle();
    	style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
    	style.setFillPattern(CellStyle.FINE_DOTS);
    	
    	//设置字体
    	Font font = workbook.createFont();
    	font.setColor(IndexedColors.YELLOW.getIndex());
    	font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    	font.setItalic(true);
//    	font.setUnderline(Font.U_DOUBLE);
    	font.setFontName("Helvetia");
    	style.setFont(font);
    	
    	cell.setCellStyle(style);
    	cell.setCellValue("我爱我家！");
    	
    	try {
			FileOutputStream output = new FileOutputStream("Test2.xls");
			workbook.write(output);
			output.close();
    	} catch (Exception e) {
			e.printStackTrace();
		}
    }
	
}
