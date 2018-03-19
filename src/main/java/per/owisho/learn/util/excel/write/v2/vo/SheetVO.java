package per.owisho.learn.util.excel.write.v2.vo;

/**
 * ExcelVO辅助类，为了存放excel表单中相关信息存在
 * @author wangyang
 * @version 1.0
 * @date 2018年01月12日
 */
public class SheetVO {
	
	/**
	 * 表单名称
	 */
	private String sheetName;
	
	/**
	 * 表单标头
	 */
	private SheetTitleVO tilte;
	
	/**
	 * 表单内容行
	 */
	private SheetContentVO content;

	public SheetVO(String sheetName, SheetTitleVO tilte, SheetContentVO content) {
		super();
		this.sheetName = sheetName;
		this.tilte = tilte;
		this.content = content;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public SheetTitleVO getTilte() {
		return tilte;
	}

	public void setTilte(SheetTitleVO tilte) {
		this.tilte = tilte;
	}

	public SheetContentVO getContent() {
		return content;
	}

	public void setContent(SheetContentVO content) {
		this.content = content;
	}

}

