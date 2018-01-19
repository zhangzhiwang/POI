package com.asiainfo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格对齐
 * 
 * @author zhangzhiwang
 * @date 2018年1月19日 下午1:23:42
 */
public class ExcelAligning {
	public static void main(String[] args) {
		OutputStream out = null;
		Workbook wb = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/ExcelAligning.xls");
			wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet();
			Row row = sheet.createRow(0);
			row.setHeightInPoints(30);//设置行高
			createCell(wb, row, 0, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
			createCell(wb, row, 1, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
			createCell(wb, row, 2, HorizontalAlignment.RIGHT, VerticalAlignment.BOTTOM);

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

	/**
	 * 
	 * 
	 * @param wb
	 * @param row
	 * @param cellIndex
	 * @param alignment
	 *            水平对齐方式
	 * @param verticalAlignment
	 *            垂直对齐方式
	 * @author zhangzhiwang
	 * @date 2018年1月19日 下午1:42:48
	 */
	private static void createCell(Workbook wb, Row row, int cellIndex, HorizontalAlignment alignment, VerticalAlignment verticalAlignment) {
		Cell cell = row.createCell(cellIndex);
		cell.setCellValue(new HSSFRichTextString("aaa"));

		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(alignment);
		cellStyle.setVerticalAlignment(verticalAlignment);
		cell.setCellStyle(cellStyle);
	}
}
