package com.asiainfo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格字体
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午3:17:45
 */
public class CellFont {
	public static void main(String[] args) {
		OutputStream out = null;
		Workbook wb = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CellFont.xls");
			wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet();
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			cell.setCellValue("中国");

			Font font = wb.createFont();
			font.setFontHeightInPoints((short) 30);
			font.setItalic(true);
			font.setFontName("宋体");
			font.setColor(IndexedColors.RED.getIndex());
			font.setStrikeout(true);//中间有条横线

			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setFont(font);
			cell.setCellStyle(cellStyle);

			wb.write(out);
			System.out.println("OK");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (out != null) {
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
