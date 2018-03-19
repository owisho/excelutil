package per.owisho.learn.util.excel.write.v1.vo;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import per.owisho.learn.util.excel.write.v1.ContentResolver;

/**
 * ExcelRowVO辅助类 @see ExcelRowVO
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class ExcelCellVO {

	/**
	 * 内容
	 */
	private String content;
	
	/**
	 * 内容样式
	 */
	private HSSFCellStyle style;

	/**
	 * 使用解析器将内容解析成字符串格式便于excel输出
	 * @param content
	 * @param style
	 * @param resolver
	 */
	public ExcelCellVO(Object content,HSSFCellStyle style,ContentResolver resolver) {
		super();
		this.style = style;
		if(content==null) {
			this.content = "";
			return ;
		}
		if(resolver==null) {
			this.content = content.toString();
		}else {
			this.content = resolver.resolve(content);
		}
	}
	
	/**
	 * 日期格式内容简化的构造方式（可以用解析器方式替代）
	 * @param content
	 * @param style
	 * @param format
	 */
	public ExcelCellVO(Date content, HSSFCellStyle style,String format) {
		super();
		this.style = style;
		if(content==null) {
			this.content = "";
			return;
		}
		if(format==null)
			format = "yyyy-MM-dd";
		this.content = new SimpleDateFormat(format).format(content);
	}
	
	/**
	 * 字符串格式内容构造方式
	 * @param content
	 * @param style
	 */
	public ExcelCellVO(String content, HSSFCellStyle style) {
		super();
		this.content = content==null?"":content;
		this.style = style;
	}

	public String getContent() {
		return content;
	}

	public void setContent(String content) {
		this.content = content;
	}

	public HSSFCellStyle getStyle() {
		return style;
	}

	public void setStyle(HSSFCellStyle style) {
		this.style = style;
	}
	
}
