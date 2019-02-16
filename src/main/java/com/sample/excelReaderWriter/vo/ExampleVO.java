package com.sample.excelReaderWriter.vo;

public class ExampleVO {
	
    private long templateId;   
    private String templateName;
	private long sheetId;
	private String sheetName;
	private int sheetIndex;
	private long headerId;
	private int colFrom;
	private int colTo;
	private int rowFrom;
	private int rowTo;
	private String headerText;
	private String headerColor;
	private String columnKey;
	
	public long getTemplateId() {
		return templateId;
	}
	public void setTemplateId(long templateId) {
		this.templateId = templateId;
	}
	public String getTemplateName() {
		return templateName;
	}
	public void setTemplateName(String templateName) {
		this.templateName = templateName;
	}
	public long getSheetId() {
		return sheetId;
	}
	public void setSheetId(long sheetId) {
		this.sheetId = sheetId;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	public int getSheetIndex() {
		return sheetIndex;
	}
	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}
	public long getHeaderId() {
		return headerId;
	}
	public void setHeaderId(long headerId) {
		this.headerId = headerId;
	}
	public int getColFrom() {
		return colFrom;
	}
	public void setColFrom(int colFrom) {
		this.colFrom = colFrom;
	}
	public int getColTo() {
		return colTo;
	}
	public void setColTo(int colTo) {
		this.colTo = colTo;
	}
	public int getRowFrom() {
		return rowFrom;
	}
	public void setRowFrom(int rowFrom) {
		this.rowFrom = rowFrom;
	}
	public int getRowTo() {
		return rowTo;
	}
	public void setRowTo(int rowTo) {
		this.rowTo = rowTo;
	}
	public String getHeaderText() {
		return headerText;
	}
	public void setHeaderText(String headerText) {
		this.headerText = headerText;
	}
	public String getHeaderColor() {
		return headerColor;
	}
	public void setHeaderColor(String headerColor) {
		this.headerColor = headerColor;
	}
	public String getColumnKey() {
		return columnKey;
	}
	public void setColumnKey(String columnKey) {
		this.columnKey = columnKey;
	}
	@Override
	public String toString() {
		return "ExampleVO [templateId=" + templateId + ", templateName=" + templateName + ", sheetId=" + sheetId
				+ ", sheetName=" + sheetName + ", sheetIndex=" + sheetIndex + ", headerId=" + headerId + ", colFrom="
				+ colFrom + ", colTo=" + colTo + ", rowFrom=" + rowFrom + ", rowTo=" + rowTo + ", headerText="
				+ headerText + ", headerColor=" + headerColor + ", columnKey=" + columnKey + "]";
	}
	
	
	

}
