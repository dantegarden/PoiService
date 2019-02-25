package com.dvt.PoiService.business.main.dto;

public class FormDTO {
	private String OdooUrl;
	private String uname;
	private String pwd;
	private String db;
	private String modelName;
	private Integer action_id;
	private Integer form_id;
	private Integer tree_id;
	private Integer search_id;
	public String getOdooUrl() {
		return OdooUrl;
	}
	public void setOdooUrl(String odooUrl) {
		OdooUrl = odooUrl;
	}
	public String getUname() {
		return uname;
	}
	public void setUname(String uname) {
		this.uname = uname;
	}
	public String getPwd() {
		return pwd;
	}
	public void setPwd(String pwd) {
		this.pwd = pwd;
	}
	public String getDb() {
		return db;
	}
	public void setDb(String db) {
		this.db = db;
	}
	public Integer getAction_id() {
		return action_id;
	}
	public void setAction_id(Integer action_id) {
		this.action_id = action_id;
	}
	public Integer getForm_id() {
		return form_id;
	}
	public void setForm_id(Integer form_id) {
		this.form_id = form_id;
	}
	public Integer getTree_id() {
		return tree_id;
	}
	public void setTree_id(Integer tree_id) {
		this.tree_id = tree_id;
	}
	public Integer getSearch_id() {
		return search_id;
	}
	public void setSearch_id(Integer search_id) {
		this.search_id = search_id;
	}
	
	public String getModelName() {
		return modelName;
	}
	public void setModelName(String modelName) {
		this.modelName = modelName;
	}
	
	public FormDTO(String odooUrl, String uname, String pwd, String db,
			String modelName, Integer action_id, Integer form_id,
			Integer tree_id, Integer search_id) {
		super();
		OdooUrl = odooUrl;
		this.uname = uname;
		this.pwd = pwd;
		this.db = db;
		this.modelName = modelName;
		this.action_id = action_id;
		this.form_id = form_id;
		this.tree_id = tree_id;
		this.search_id = search_id;
	}
	public FormDTO() {
		super();
	}
	
	
}
