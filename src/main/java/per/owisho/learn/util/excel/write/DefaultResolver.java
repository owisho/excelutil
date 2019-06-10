package per.owisho.learn.util.excel.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 默认内容处理器类
 * @author owisho
 * @date 2018-03-19
 */
public class DefaultResolver{

	/**
	 * 默认标题解析器
	 */
	private static Title titleResolver = null;
	/**
	 * 默认内容解析器
	 */
	private static Content contentResolver = null;


	public static ContentResolver getTitleResolver(){
		if(titleResolver==null){
			titleResolver=new Title();
		}
		return titleResolver;
	}

	public static ContentResolver getContentResolver(){
		if(contentResolver==null){
			contentResolver=new Content();
		}
		return contentResolver;
	}

}

/**
 * 默认excel表头处理器
 * @author owisho
 *
 */
class Title implements ContentResolver{

	private CellStyle cellStyle;

	@Override
	public void resolve(Cell cell, Object content) {
		cell.setCellValue(content.toString());
		Sheet sheet = cell.getSheet();
		if(cellStyle==null){
			cellStyle = this.getTitleStyle(sheet.getWorkbook());
		}
		cell.setCellStyle(cellStyle);
	}
	
}

/**
 * 默认excel单元格内容处理器
 * @author owisho
 */
class Content implements ContentResolver{

	private CellStyle cellStyle;

	@Override
	public void resolve(Cell cell, Object content) {
		cell.setCellValue(content.toString());
		Sheet sheet = cell.getSheet();
		if(cellStyle==null){
			System.out.println(this);
			System.out.println("重新获取style");
			cellStyle = this.getContentStyle(sheet.getWorkbook());
		}
		cell.setCellStyle(cellStyle);
//		if(sheet instanceof SXSSFSheet){
//			((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
//		}
//		sheet.autoSizeColumn(cell.getColumnIndex());
	}

}
