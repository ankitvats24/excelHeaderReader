package com.sample.excelReaderWriter.dao;

import com.sample.excelReaderWriter.entity.TemplateMasterTest;

public interface HeaderDao {
	
	public void createHeader(TemplateMasterTest templateMasterTest);
	public TemplateMasterTest readHeader(long templateId);
	public long getTemplateId(String templateName);
	public void deleteTemplate(long templateId);

}
