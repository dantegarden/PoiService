package com.dvt.PoiService.business.main.dto;

import java.io.File;

import org.springframework.web.multipart.MultipartFile;

public class UploadDTO {
	
	//多文件上传的文件对象，多文件是一个一个文件上传，所以multiFile是当前正在上传的文件对象
	private MultipartFile multiFile;
	
    //多文件上传文件对象的文名，当前正在上传的文件名
    private String multiFileName;
    
    //属性值，接收webupload自带的参数
    private String chunk; // 当前第几个分片
    private String chunks;// 总分片个数
    private String size;// 单个文件的总大小

    //自定义属性值
    private String fileSize;//所有文件的总大小
    private String[] fileName;// 文件名列表
    private String[] fileSizeOneByOne;//每个文件大小
    private String status;
	
	
	
	public MultipartFile getMultiFile() {
		return multiFile;
	}
	public void setMultiFile(MultipartFile multiFile) {
		this.multiFile = multiFile;
	}
	public String getMultiFileName() {
		return multiFileName;
	}
	public void setMultiFileName(String multiFileName) {
		this.multiFileName = multiFileName;
	}
	public String[] getFileName() {
		return fileName;
	}
	public void setFileName(String[] fileName) {
		this.fileName = fileName;
	}
	public String getChunk() {
		return chunk;
	}
	public void setChunk(String chunk) {
		this.chunk = chunk;
	}
	public String getChunks() {
		return chunks;
	}
	public void setChunks(String chunks) {
		this.chunks = chunks;
	}
	public String getSize() {
		return size;
	}
	public void setSize(String size) {
		this.size = size;
	}
	public String getFileSize() {
		return fileSize;
	}
	public void setFileSize(String fileSize) {
		this.fileSize = fileSize;
	}
	public String[] getFileSizeOneByOne() {
		return fileSizeOneByOne;
	}
	public void setFileSizeOneByOne(String[] fileSizeOneByOne) {
		this.fileSizeOneByOne = fileSizeOneByOne;
	}
	public String getStatus() {
		return status;
	}
	public void setStatus(String status) {
		this.status = status;
	}
	
	public UploadDTO() {
		super();
	}
	public UploadDTO(MultipartFile multiFile, String multiFileName,
			String chunk, String chunks, String size, String fileSize,
			String[] fileName, String[] fileSizeOneByOne, String status) {
		super();
		this.multiFile = multiFile;
		this.multiFileName = multiFileName;
		this.chunk = chunk;
		this.chunks = chunks;
		this.size = size;
		this.fileSize = fileSize;
		this.fileName = fileName;
		this.fileSizeOneByOne = fileSizeOneByOne;
		this.status = status;
	}
	
    
    
}
