package com.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * 利用组件POI3.0.2动态导出EXCEL文档
 */
public class ExportExcel<T> {

	/**
	 * @param sheetName sheet名字
	 * @param headers 表格属�?列名数组
	 * @param map �?��显示的数据集�?         
	 * @param filedNames 对应列明的实体字段名�?
	 * @param out 与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
	 */
	@SuppressWarnings({ "unchecked", "deprecation" })
	public void exportExcel(String sheetName, String[] headers, List<Map<String, Object>> map, String[] filedNames, OutputStream out) {
		// 声明工作
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 生成表格
		HSSFSheet sheet = workbook.createSheet(sheetName);
		// 设置表格默认列宽度为15个字
		sheet.setDefaultColumnWidth((short) 15);
		// 生成样式
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置这些样式
		style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		// 生成字体
		HSSFFont font = workbook.createFont();
		font.setColor(HSSFColor.VIOLET.index);
		font.setFontHeightInPoints((short) 12);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 把字体应用到当前的样式
		style.setFont(font);
		// 生成并设置excel样式
		HSSFCellStyle style2 = workbook.createCellStyle();
		style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
		style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		// 生成另一个字
		HSSFFont font2 = workbook.createFont();
		font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		// 把字体应用到当前的样
		style2.setFont(font2);

		HSSFFont font3 = workbook.createFont();
		font3.setColor(HSSFColor.BLUE.index);

		// 产生表格标题
		HSSFRow row = sheet.createRow(0);
		for (short i = 0; i < headers.length; i++) {
			HSSFCell cell = row.createCell(i);
			cell.setCellStyle(style);
			HSSFRichTextString text = new HSSFRichTextString(headers[i]);
			cell.setCellValue(text);
		}
		// 遍历集合数据，产生数据行
		List<Map<String, Object>> it = map;
		int index = 0;
		for (Map<String,Object> itMap : it) {
			index++;
			row = sheet.createRow(index);
			int i = 0;
			for (String item : filedNames) {
				HSSFCell cell = row.createCell(i);
				cell.setCellStyle(style2);
				HSSFRichTextString richString = new HSSFRichTextString(itMap.get(item) == null ? "" : itMap.get(item).toString());
				richString.applyFont(font3);
				cell.setCellValue(richString.getString());
				i++;
			}
		}
		try {
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}