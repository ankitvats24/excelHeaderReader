package com.sample.excelReaderWriter.entity;

import java.util.HashSet;
import java.util.Set;

import javax.persistence.CascadeType;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.JoinTable;
import javax.persistence.OneToMany;

@Entity
public class SheetMasterTest {

	@Id
	@GeneratedValue(strategy=GenerationType.AUTO)
	private long sheetId;
	private long templateId;
	private String sheetName;
	private int sheetIndex;
	private Set<SheetHeaderMasterTest> sheetHeaderMasterTestList = new HashSet<SheetHeaderMasterTest>(0);
	
	public long getSheetId() {
		return sheetId;
	}
	public void setSheetId(long sheetId) {
		this.sheetId = sheetId;
	}
	public long getTemplateId() {
		return templateId;
	}
	public void setTemplateId(long templateId) {
		this.templateId = templateId;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	@OneToMany(cascade = CascadeType.ALL)
	@JoinTable(name = "Sheet_Header_Master_Test", joinColumns = { @JoinColumn(name = "SHEET_ID") })
	public Set<SheetHeaderMasterTest> getSheetHeaderMasterTestList() {
		return this.sheetHeaderMasterTestList;
	}
	public void setSheetHeaderMasterTestList(Set<SheetHeaderMasterTest> sheetHeaderMasterTestList) {
		this.sheetHeaderMasterTestList = sheetHeaderMasterTestList;
	}
	public int getSheetIndex() {
		return sheetIndex;
	}
	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}
	@Override
	public String toString() {
		return "SheetMasterTest [sheetId=" + sheetId + ", templateId=" + templateId + ", sheetName=" + sheetName
				+ ", sheetIndex=" + sheetIndex + ", sheetHeaderMasterTestList=" + sheetHeaderMasterTestList + "]";
	}
	

}
