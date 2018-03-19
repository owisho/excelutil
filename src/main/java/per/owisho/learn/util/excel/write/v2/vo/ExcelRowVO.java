package per.owisho.learn.util.excel.write.v2.vo;

import java.util.List;

/**
 * SheetTitleVO辅助类及SheetContentVO辅助类 @see SheetTitleVO 和 @see SheetContentVO
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class ExcelRowVO {

	/**
	 * 行内元素
	 */
	private List<ExcelCellVO> cells;

	public ExcelRowVO(List<ExcelCellVO> cells) {
		super();
		this.cells = cells;
	}

	public List<ExcelCellVO> getCells() {
		return cells;
	}

	public void setCells(List<ExcelCellVO> cells) {
		this.cells = cells;
	}
	
}
