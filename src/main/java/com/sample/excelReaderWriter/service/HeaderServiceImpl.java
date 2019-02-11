package com.sample.excelReaderWriter.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.sample.excelReaderWriter.dao.HeaderDao;
import com.sample.excelReaderWriter.entity.TemplateMasterTest;
@Service

public class HeaderServiceImpl implements HeaderService {

	@Autowired
	HeaderDao headerDao;

	@Override
	public void createHeader(TemplateMasterTest templateMasterTest) {
		// TODO Auto-generated method stub

	}

	@Override
	public TemplateMasterTest readHeader(long templateId) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public long getTemplateId(String templateName) {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public void deleteTemplate(long templateId) {
		// TODO Auto-generated method stub

	}

}
