package com.dvt.PoiService.business.example.webservice;

import javax.jws.WebService;
/**
* 接口定义
*  @WebService
*  用于定义webservice对外开放的接口
*/
@WebService
public interface MyWebService {
	public String SayHello(String name);
}
