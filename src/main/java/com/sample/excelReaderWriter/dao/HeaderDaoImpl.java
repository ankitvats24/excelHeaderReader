package com.sample.excelReaderWriter.dao;

import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import com.sample.excelReaderWriter.entity.TemplateMasterTest;

@Repository
public class HeaderDaoImpl implements HeaderDao {
	
    @Autowired
    private SessionFactory sessionFactory;

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
