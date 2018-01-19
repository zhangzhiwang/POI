package com.asiainfo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格边框
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午2:14:03
 */
public class CellBorder {
	public static void main(String[] args) {
		OutputStream out = null;
		Workbook wb = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CellBorder.xls");
			wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet();
			Row row = sheet.createRow(1);
			Cell cell = row.createCell(1);
			cell.setCellValue("aaa");
			CellStyle cellStyle = wb.createCellStyle();
			//上边框
			cellStyle.setBorderTop(BorderStyle.THIN);//设置粗细
			cellStyle.setTopBorderColor(IndexedColors.RED.getIndex());//设置颜色

			//左边框
			cellStyle.setBorderLeft(BorderStyle.MEDIUM);
			cellStyle.setLeftBorderColor(IndexedColors.BLUE.getIndex());
			
			//右边框
			cellStyle.setBorderRight(BorderStyle.DOUBLE);
			cellStyle.setRightBorderColor(IndexedColors.GREEN.getIndex());
			
			//底边框
			cellStyle.setBorderBottom(BorderStyle.DASHED);
			cellStyle.setBottomBorderColor(IndexedColors.BROWN.getIndex());
			
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
