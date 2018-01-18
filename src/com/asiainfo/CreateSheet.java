package com.asiainfo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 创建工作表
 * 
 * @author zhangzhiwang
 * @date 2017年11月14日 下午9:20:30
 */
public class CreateSheet {
	public static void main(String[] args) {
		// 创建工作表前先创建工作簿
		Workbook workbook = new HSSFWorkbook();
		workbook.createSheet("创建的第一个sheet");
		workbook.createSheet("创建的第二个sheet");
		workbook.createSheet();
		OutputStream out = null;
		try {
			out = new FileOutputStream("/Users/zhangzhiwang/Documents/poi/CreateSheet.xls");
			workbook.write(out);
			System.out.println("OK!");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (out != null) {
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
