package com.asiainfo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格内部换行以及自定义数据格式
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午4:01:08
 */
public class CellLineFeed {
	public static void main(String[] args) {
		OutputStream out = null;
		Workbook wb = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CellLineFeed.xls");
			wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet();
			Row row0 = sheet.createRow(0);
			row0.setHeightInPoints(sheet.getDefaultRowHeightInPoints() * 3);
			Cell cell0 = row0.createCell(0);
			cell0.setCellValue("中国\n北京");

			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setWrapText(true);// 开启单元格内换行
			cell0.setCellStyle(cellStyle);
			
			Cell cell1 = row0.createCell(1);
			cell1.setCellValue(1234567890.19);
			//自定义数据格式
			CellStyle cellStyle2 = wb.createCellStyle();
			DataFormat dataFormat = wb.createDataFormat();
//			cellStyle2.setDataFormat(dataFormat.getFormat("0.0"));//保留一位小数
			cellStyle2.setDataFormat(dataFormat.getFormat("#,##0.00"));//每三位数加一个逗号
			cell1.setCellStyle(cellStyle2);
			
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
