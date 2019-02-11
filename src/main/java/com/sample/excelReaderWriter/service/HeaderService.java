package com.sample.excelReaderWriter.service;

import com.sample.excelReaderWriter.entity.TemplateMasterTest;

public interface HeaderService {

	public void createHeader(TemplateMasterTest templateMasterTest);
	public TemplateMasterTest readHeader(long templateId);
	public long getTemplateId(String templateName);
	public void deleteTemplate(long templateId);
}
