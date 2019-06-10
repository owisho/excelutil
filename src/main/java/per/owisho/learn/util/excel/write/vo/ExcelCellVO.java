package per.owisho.learn.util.excel.write.vo;

import org.apache.poi.xssf.streaming.SXSSFCell;
import per.owisho.learn.util.excel.write.ContentResolver;


/**
 * ExcelRowVO辅助类 @see ExcelRowVO
 * @author wangyang
 * @version 2.0
 * @date 2018年01月12日
 */
public class ExcelCellVO {

	/**
	 * 内容
	 */
	private Object content;
	
	/**
	 * 内容处理方式
	 */
	private ContentResolver resolver;
	
	/**
	 * 使用解析器将内容解析成字符串格式便于excel输出
	 * @param content
	 * @param resolver
	 */
	public ExcelCellVO(Object content,ContentResolver resolver) {
		super();
		this.content = content==null?"":content;
		this.resolver = resolver;
	}

	public void drawCell(SXSSFCell cell) {
		try {
			resolver.resolve(cell, this.content);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
