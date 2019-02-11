package com.sample.excelReaderWriter.entity;

import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;

@Entity
public class SheetHeaderMasterTest {

	@Id
	@GeneratedValue(strategy=GenerationType.AUTO)
	private long headerId;
	private long sheetId;
	private int colFrom;
	private int colTo;
	private int rowFrom;
	private int rowTo;
	private String headerText;
	
	public long getHeaderId() {
		return headerId;
	}
	public void setHeaderId(long headerId) {
		this.headerId = headerId;
	}
	public long getSheetId() {
		return sheetId;
	}
	public void setSheetId(long sheetId) {
		this.sheetId = sheetId;
	}
	public String getHeaderText() {
		return headerText;
	}
	public void setHeaderText(String headerText) {
		this.headerText = headerText;
	}
	@Override
	public String toString() {
		return "SheetHeaderMasterTest \n[headerId=" + headerId + ", sheetId=" + sheetId + ", colFrom=" + colFrom
				+ ", colTo=" + colTo + ", rowFrom=" + rowFrom + ", rowTo=" + rowTo + ", headerText=" + headerText + "]";
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
	
}
