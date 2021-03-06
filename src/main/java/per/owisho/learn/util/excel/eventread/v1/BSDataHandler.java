package per.owisho.learn.util.excel.eventread.v1;

import java.util.List;

/**
 * 业务数据处理接口
 * @author owisho
 * @version 1.0
 * @date 2018年1月19日
 */
public interface BSDataHandler {
	
	/**
	 * 业务数据处理接口，用来处理业务数据，接口禁止异常抛出，
	 * 异常信息应通过接口实现类使用其他方式记录，
	 * 并在调用excel解析的方法中处理
	 * @param datas
	 */
	@SuppressWarnings("rawtypes")
	void process(List[] datas);
	
}
