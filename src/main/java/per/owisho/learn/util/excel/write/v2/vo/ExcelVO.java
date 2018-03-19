package per.owisho.learn.util.excel.write.v2.vo;

import java.util.List;

/**
 * 用于封装输出实体并转化成excel文档
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class ExcelVO {

	/**
	 * excel名称
	 */
	private String excelName;
	
	/**
	 * excel表单
	 */
	private List<SheetVO> sheets;

	public ExcelVO(String excelName, List<SheetVO> sheets) {
		super();
		this.excelName = excelName;
		this.sheets = sheets;
	}

	public String getExcelName() {
		return excelName;
	}

	public void setExcelName(String excelName) {
		this.excelName = excelName;
	}

	public List<SheetVO> getSheets() {
		return sheets;
	}

	public void setSheets(List<SheetVO> sheets) {
		this.sheets = sheets;
	}
	
}


