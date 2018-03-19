package per.owisho.learn.util.excel.write.v2;

import org.apache.poi.hssf.usermodel.HSSFCell;

/**
 * 默认内容处理器类
 * @author owisho
 * @date 2018-03-19
 */
public enum DefaultResolverEnum {

	Content(new Content()),Title(new Title());
	
	private ContentResolver resolver;
	
	private DefaultResolverEnum(ContentResolver resolver) {
		this.resolver = resolver;
	}
	
	public ContentResolver getResolver() {
		return this.resolver;
	}
	
}

/**
 * 默认excel表头处理器
 * @author owisho
 *
 */
class Title implements ContentResolver{

	@Override
	public void resolve(HSSFCell cell, Object content) {
		cell.setCellValue(content.toString());
		cell.setCellStyle(this.getTitleStyle(cell.getSheet().getWorkbook()));
	}
	
}

/**
 * 默认excel单元格内容处理器
 * @author owisho
 */
class Content implements ContentResolver{

	@Override
	public void resolve(HSSFCell cell, Object content) {
		cell.setCellValue(content.toString());
		cell.setCellStyle(this.getContentStyle(cell.getSheet().getWorkbook()));
	}

}
