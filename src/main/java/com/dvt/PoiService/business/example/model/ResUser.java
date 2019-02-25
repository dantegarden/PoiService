package com.dvt.PoiService.business.example.model;

import java.sql.Timestamp;
import java.util.Date;

import com.dvt.PoiService.commons.entity.BaseBean;
import com.dvt.PoiService.commons.entity.BaseEntity;

public class ResUser extends BaseBean{
	private Integer id;
	private String login;
	private Timestamp create_date;
	
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public String getLogin() {
		return login;
	}
	public void setLogin(String login) {
		this.login = login;
	}
	public Date getCreate_date() {
		return create_date;
	}
	public void setCreate_date(Timestamp create_date) {
		this.create_date = create_date;
	}
	public ResUser(Integer id, String login, Timestamp create_date) {
		super();
		this.id = id;
		this.login = login;
		this.create_date = create_date;
	}
	public ResUser() {
		super();
	}
	
	
}
