package com.dvt.PoiService.business.main.dto;

import java.util.List;

public class ExcelDTO {
	private String sourcePath;
	private String targetPath;
	private List<SheetDTO> sheetList;
	public ExcelDTO(String sourcePath, String targetPath, List<SheetDTO> sheetList) {
		super();
		this.sourcePath = sourcePath;
		this.targetPath = targetPath;
		this.sheetList = sheetList;
	}
	public String getSourcePath() {
		return sourcePath;
	}
	public void setSourcePath(String sourcePath) {
		this.sourcePath = sourcePath;
	}
	public String getTargetPath() {
		return targetPath;
	}
	public void setTargetPath(String targetPath) {
		this.targetPath = targetPath;
	}
	public List<SheetDTO> getSheetList() {
		return sheetList;
	}
	public void setSheetList(List<SheetDTO> sheetList) {
		this.sheetList = sheetList;
	}
	
	
}
