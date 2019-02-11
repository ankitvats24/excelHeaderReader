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

public class TemplateMasterTest {

    @Id
    @GeneratedValue(strategy=GenerationType.AUTO)
    private long templateId;   
    private String templateName;
    private Set<SheetMasterTest> sheetMasterTestList = new HashSet<SheetMasterTest>(0);
    
    public TemplateMasterTest() {
	}
	public TemplateMasterTest(long templateId, String templateName, Set<SheetMasterTest> sheetMasterTestList) {
		super();
		this.templateId = templateId;
		this.templateName = templateName;
		this.sheetMasterTestList = sheetMasterTestList;
	}

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
	@OneToMany(cascade = CascadeType.ALL)
	@JoinTable(name = "Sheet_Master_Test", joinColumns = { @JoinColumn(name = "TEMPLATE_ID") })
	public Set<SheetMasterTest> getSheetMasterTestList() {
		return this.sheetMasterTestList;
	}

	public void setSheetMasterTestList(Set<SheetMasterTest> sheetMasterTestList) {
		this.sheetMasterTestList = sheetMasterTestList;
	}
	@Override
	public String toString() {
		return "TemplateMasterTest [templateId=" + templateId + ", templateName=" + templateName
				+ ", sheetMasterTestList=" + sheetMasterTestList + "]";
	}
 
}