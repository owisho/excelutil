package per.owisho.learn.poi.read;

import java.util.List;

public interface DataHandler {
	
	/**
	 * 处理类需要自己处理在数据处理中的异常信息
	 * @param datas
	 */
	@SuppressWarnings("rawtypes")
	void process(List[] datas);
	
}
