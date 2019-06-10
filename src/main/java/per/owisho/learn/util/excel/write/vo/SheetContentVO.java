package per.owisho.learn.util.excel.write.vo;

import java.util.List;

/**
 * 表单内容 @see SheetVO 辅助类
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class SheetContentVO {
	
	/**
	 * 内容行
	 */
	private List<ExcelRowVO> rows;

	public SheetContentVO(List<ExcelRowVO> rows) {
		super();
		this.rows = rows;
	}

	public List<ExcelRowVO> getRows() {
		return rows;
	}

	public void setRows(List<ExcelRowVO> rows) {
		this.rows = rows;
	}
	
}
