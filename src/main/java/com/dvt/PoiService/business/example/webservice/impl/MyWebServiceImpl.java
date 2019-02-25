package com.dvt.PoiService.business.example.webservice.impl;

import com.dvt.PoiService.business.example.webservice.MyWebService;

public class MyWebServiceImpl implements MyWebService {

	@Override
	public String SayHello(String name) {
		return "HelloWorld!! " + name;
	}
	
}
