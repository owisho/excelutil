package per.owisho.learn.util.excel.write;

/**
 * excel单元格内容处理器
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public interface ContentResolver {

	/**
	 * 将对象类型数据转化成要求的字符串类型数据
	 * @param content
	 * @return
	 */
	String resolve(Object content); 
	
}
