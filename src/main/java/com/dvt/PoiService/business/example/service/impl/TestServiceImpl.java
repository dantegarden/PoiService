package com.dvt.PoiService.business.example.service.impl;

import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.dvt.PoiService.business.example.model.ResUser;
import com.dvt.PoiService.business.example.service.TestService;
@Service
public class TestServiceImpl implements TestService{
	
	private static final Logger logger = LoggerFactory.getLogger(TestServiceImpl.class);

	@Override
	public ResUser getResUsers() {
		return null;
	}

	
	
}
