package com.dvt.PoiService.business.main.dto;

import java.util.List;

public class RowDTO {
	private int startRow;//现在的起始行
	private int insertRow;//插入多少行
	private List<List<String>> myRows;//插入的行的值
	private String type;//操作类型
	private boolean writeFromStart = Boolean.FALSE;//是否从起始行开始
	private int orignStartRow;//模板的起始行
	
	private int[] sourceRows;
	private int targetStartRow;
	
	public int getStartRow() {
		return startRow;
	}
	public void setStartRow(int startRow) {
		this.startRow = startRow;
		this.orignStartRow = startRow;
	}
	
	public int getInsertRow() {
		return insertRow;
	}
	public void setInsertRow(int insertRow) {
		this.insertRow = insertRow;
	}

	public List<List<String>> getMyRows() {
		return myRows;
	}
	public void setMyRows(List<List<String>> myRows) {
		this.myRows = myRows;
	}
	
	public RowDTO(int startRow,
			List<List<String>> myRows) {
		super();
		this.startRow = startRow;
		this.orignStartRow = startRow;
		this.myRows = myRows;
	}
	
	public RowDTO(int startRow, int insertRow) {
		super();
		this.startRow = startRow;
		this.orignStartRow = startRow;
		this.insertRow = insertRow;
	}
	
	public RowDTO(int[] sourceRows, int targetStartRow) {
		super();
		this.type = "copy";
		this.sourceRows = sourceRows;
		this.targetStartRow = targetStartRow;
	}
	
	public RowDTO(int startRow, List<List<String>> myRows,
			 boolean writeFromStart) {
		super();
		this.startRow = startRow;
		this.orignStartRow = startRow;
		this.myRows = myRows;
		this.writeFromStart = writeFromStart;
	}
	public RowDTO(int startRow, List<List<String>> myRows,
			String type, boolean writeFromStart) {
		super();
		this.startRow = startRow;
		this.orignStartRow = startRow;
		this.myRows = myRows;
		this.type = type;
		this.writeFromStart = writeFromStart;
	}
	public RowDTO(int startRow, List<List<String>> myRows,
			String type) {
		super();
		this.startRow = startRow;
		this.orignStartRow = startRow;
		this.myRows = myRows;
		this.type = type;
	}
	public void shiftStartRow(int shift){
		this.startRow += shift;
	}
	public boolean isWriteFromStart() {
		return writeFromStart;
	}
	public void setWriteFromStart(boolean writeFromStart) {
		this.writeFromStart = writeFromStart;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public int getOrignStartRow() {
		return orignStartRow;
	}
	public int[] getSourceRows() {
		return sourceRows;
	}
	public void setSourceRows(int[] sourceRows) {
		this.sourceRows = sourceRows;
	}
	public int getTargetStartRow() {
		return targetStartRow;
	}
	public void setTargetStartRow(int targetStartRow) {
		this.targetStartRow = targetStartRow;
	}
	
	
	
}
