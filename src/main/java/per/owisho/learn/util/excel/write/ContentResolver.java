package per.owisho.learn.util.excel.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

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
	void resolve(Cell cell, Object content);

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
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
		cellStyle.setAlignment(XSSFCellStyle.VERTICAL_BOTTOM );
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
