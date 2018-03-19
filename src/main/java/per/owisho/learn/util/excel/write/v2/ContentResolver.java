package per.owisho.learn.util.excel.write.v2;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * excel单元格内容处理器 为了支持在excel单元格中画图片
 * 
 * @author wangyang
 * @version 2.0
 * @date 2018年01月12日
 */
public interface ContentResolver {

	/**
	 * 数据内容进行处理
	 * 
	 * @param content
	 * @return
	 */
	void resolve(HSSFCell cell, Object content);

	default CellStyle getContentStyle(Workbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		Font ztFont = wb.createFont();
		ztFont.setFontHeightInPoints((short) 12);
		ztFont.setFontName("宋体");
		cellStyle.setFont(ztFont);
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		return cellStyle;
	}

	default CellStyle getTitleStyle(Workbook wb) {
		CellStyle cellStyle = wb.createCellStyle();
		Font ztFont = wb.createFont();
		ztFont.setFontHeightInPoints((short) 14);
		ztFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cellStyle.setFont(ztFont);
		cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		return cellStyle;
	}

}
