package com.asiainfo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 创建单元格
 * 
 * @author zhangzhiwang
 * @date 2017年11月14日 下午9:37:00
 */
public class CreateCell {
	public static void main(String[] args) {
		//创建单元格的顺序是：先创建工作簿，然后创建工作表，然后创建行，最后创建单元格。注：POI没有创建列的概念，创建完行之后创建的第几个单元格就相当于该行的第几列
		//创建工作簿
		Workbook workbook = new HSSFWorkbook();
		//创建工作表
		Sheet sheet1 = workbook.createSheet("第一个工作表");//注：行和单元格是按照索引来创建的，工作表没有按照索引来创建的api
		//创建第一行
		Row row1 = sheet1.createRow(0);
		//创建第一行的第一个单元格（列）
		Cell row1Cell1 = row1.createCell(0);
		row1Cell1.setCellValue(123);//给单元格设置值
		
		//创建第二个工作表
		Sheet sheet2 = workbook.createSheet("第二个工作表");
		//创建第二张工作表的第二行
		Row row2 = sheet2.createRow(1);
		//创建第二行的第三个单元格
		Cell row2Cell3 = row2.createCell(2);
		row2Cell3.setCellValue(true);
		
		//输出工作簿
		OutputStream out = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CreateCell.xls");
			workbook.write(out);
			System.out.println("OK!");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if(out != null) {
				try {
					out.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
	}
}
