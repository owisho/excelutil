package per.owisho.learn.util.excel.write.v2.vo;

import org.apache.poi.hssf.usermodel.HSSFCell;

import per.owisho.learn.util.excel.write.v2.ContentResolver;

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
	 * @param style
	 * @param resolver
	 */
	public ExcelCellVO(Object content,ContentResolver resolver) {
		super();
		if(content==null) {
			this.content = "";
			return ;
		}
		this.content = content;
		this.resolver = resolver;
	}
	
	/**
	 * 内容构造方式
	 * @param content
	 * @param style
	 */
	public ExcelCellVO(Object content) {
		super();
		this.content = content==null?"":content;
	}

	public void drawCell(HSSFCell cell) {
		resolver.resolve(cell, this.content);
	}
	
}
