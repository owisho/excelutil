package per.owisho.learn.util.excel.write.vo;

import java.util.List;

/**
 * sheetVO辅助类@see SheetVO
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class SheetTitleVO {

	/**
	 * 内容行
	 */
	private List<ExcelRowVO> rows;

	public SheetTitleVO(List<ExcelRowVO> rows) {
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
