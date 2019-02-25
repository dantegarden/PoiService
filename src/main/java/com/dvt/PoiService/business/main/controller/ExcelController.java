package com.dvt.PoiService.business.main.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.Base64;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.alibaba.druid.util.jdbc.ResultSetMetaDataBase;
import com.dvt.PoiService.business.main.dto.RowDTO;
import com.dvt.PoiService.business.main.dto.FormDTO;
import com.dvt.PoiService.business.main.dto.ExcelDTO;
import com.dvt.PoiService.business.main.dto.SheetDTO;
import com.dvt.PoiService.commons.entity.Result;
import com.dvt.PoiService.commons.utils.Base64Utils;
import com.dvt.PoiService.commons.utils.CommonHelper;
import com.dvt.PoiService.commons.utils.FileUtils;
import com.dvt.PoiService.commons.utils.JsonUtils;
import com.dvt.PoiService.commons.utils.POIConvertHtml2Excel;
import com.dvt.PoiService.commons.utils.POIReadExcelToHtml;
import com.dvt.PoiService.commons.utils.PoiUtils;
import com.dvt.PoiService.commons.utils.XmlRpcUtils;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Lists;
import com.sun.tools.internal.ws.processor.model.Request;
@Controller
@RequestMapping("/excel")
public class ExcelController {
	private static final Logger logger = LoggerFactory.getLogger(ExcelController.class);
	private static final String LINE_SEP = System.getProperty("line.separator");
	private static final String PATH_SEP = File.separator;
	
	/**
	 * 接收文件流
	 * 返回文件流
	 * **/
	@RequestMapping(value="/rewrite", method = RequestMethod.POST)
	@ResponseBody
	public ResponseEntity<byte[]> rewrite(@RequestParam("file") MultipartFile tmpFile,
			@RequestParam(value="ewmJson",required=false) String ewmJson, 
			HttpServletRequest request, HttpServletResponse response) throws IOException{
		System.out.println(ewmJson);
		ResponseEntity<byte[]> result = null;
		if (tmpFile != null) {
			String tmpFileName = tmpFile.getOriginalFilename();
			String sourceDirectory = request.getSession().getServletContext().getRealPath("/uploads");
			String targetDirectory = request.getSession().getServletContext().getRealPath("/downloads");
			int dot = tmpFileName.lastIndexOf('.');
	        String ext = "";  //文件后缀名
	        if ((dot > -1) && (dot < (tmpFileName.length() - 1))) {
	            ext = tmpFileName.substring(dot + 1);
	        }
	        
	        
	        if ("xlsx".equalsIgnoreCase(ext)) {
	        	// 重命名上传的文件名
                String sourceFileName = CommonHelper.renameFileName(tmpFileName);
                String targetFileName = CommonHelper.renameFileName(tmpFileName);
                // 保存的新文件
                File sourceFile = new File(sourceDirectory, sourceFileName);
                File targetFile = new File(targetDirectory, targetFileName);
                
                try {
                    // 保存文件
                    FileUtils.copyInputStreamToFile(tmpFile.getInputStream(), sourceFile);
                    
                    // Excel处理
                    Result r = this.insertAndWrite(sourceFile,targetFile, ewmJson);
                    if(r.getSuccess()){
                    	String filename = targetFile.getName();
            	        HttpHeaders headers = new HttpHeaders();  
            	        headers.setContentDispositionFormData("fileName", filename);  
            	        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            	        result = new ResponseEntity<byte[]>(FileUtils.readAsByteArray(targetFile), headers, HttpStatus.OK);  
                    }else{
                    	System.out.println(r.getMsg());
                    	result = new ResponseEntity<byte[]>(r.getMsg().getBytes(), HttpStatus.INTERNAL_SERVER_ERROR);
                    }
                    
                } catch (IOException e) {
                    e.printStackTrace();
                } finally{
                	if(sourceFile.exists()){
                		System.out.println("source-size:"+FileUtils.GetFileSize(sourceFile));
                		sourceFile.delete();
                	}
                	if(targetFile.exists()){
                		System.out.println("target-size:"+FileUtils.GetFileSize(targetFile));
                		targetFile.delete();
                	}
                }
	        }
	        
		}else{
			result = new ResponseEntity<byte[]>("接受文件失败".getBytes(), HttpStatus.OK);
		}
		return result;
	}
	
	@RequestMapping(value="/rewrite2Html", method = RequestMethod.POST)
	@ResponseBody
	public Result rewrite2Html(@RequestParam("file") MultipartFile tmpFile,
			@RequestParam(value="ewmJson",required=false) String ewmJson, 
			HttpServletRequest request, HttpServletResponse response) throws IOException{
		System.out.println(ewmJson);
		List<String> results = Lists.newArrayList();
		if (tmpFile != null) {
			String tmpFileName = tmpFile.getOriginalFilename();
			String sourceDirectory = request.getSession().getServletContext().getRealPath("/uploads");
			int dot = tmpFileName.lastIndexOf('.');
	        String ext = "";  //文件后缀名
	        if ((dot > -1) && (dot < (tmpFileName.length() - 1))) {
	            ext = tmpFileName.substring(dot + 1);
	        }
	        
	        if ("xlsx".equalsIgnoreCase(ext)) {
	        	// 重命名上传的文件名
                String sourceFileName = CommonHelper.renameFileName(tmpFileName);
                // 保存的新文件
                File sourceFile = new File(sourceDirectory, sourceFileName);
                try {
                    // 保存文件
                    FileUtils.copyInputStreamToFile(tmpFile.getInputStream(), sourceFile);
                    ExcelDTO excel = JsonUtils.jsonToJavaBean(ewmJson, ExcelDTO.class);
                    Workbook wb = PoiUtils.getWeebWork(sourceFile.getAbsolutePath());
    				XSSFWorkbook xwb = (XSSFWorkbook) wb;
    				for (SheetDTO sheet : excel.getSheetList()) {
    					for (int s = 0; s < xwb.getNumberOfSheets(); s++) {  
    						XSSFSheet _sheet = xwb.getSheetAt(s);  
    						if(s+1 == sheet.getSheetNum()){//找到sheet
    							// Excel转HTML
    		                    String result = POIReadExcelToHtml.readExcelToHtml(xwb, _sheet, sheet.getMaintainStyle());
    		                    results.add(result);
    						}
    					}
    				}
    				return new Result(Boolean.TRUE, "转换成功", results);
                    
                }catch(Exception e){
                	e.printStackTrace();
                	return new Result(Boolean.FALSE, e.getMessage(), null);
                }finally{
                	if(sourceFile.exists()){
                		System.out.println("source-size:"+FileUtils.GetFileSize(sourceFile));
                		sourceFile.delete();
                	}
                }
	        }else{
	        	return new Result(Boolean.FALSE, "不支持xlsx以外的文件格式", null);
	        }
		}else{
			return new Result(Boolean.FALSE, "接收文件流失败", null);
		}
		
	}
	
	@CrossOrigin(maxAge = 3600)
	@RequestMapping(value="/html2Excel", method = RequestMethod.POST)
	@ResponseBody
	public Result html2Excel(@RequestParam("htmlCode") String htmlCode,
			@RequestParam(value="sheetName", required=false) String sheetName,
			HttpServletRequest request, HttpServletResponse response) throws IOException{
		
		if(StringUtils.isNotBlank(htmlCode)){
			String sourceDirectory = request.getSession().getServletContext().getRealPath("/uploads");
			String sourceFileName = CommonHelper.renameFileName("a.xlsx");
			File sourceFile = new File(sourceDirectory, sourceFileName);
			FileOutputStream fileOutputStream = null;
			try {
				XSSFWorkbook wb = POIConvertHtml2Excel.table2Excel(htmlCode, sheetName);
				if(!sourceFile.exists()){
					sourceFile.createNewFile();
				}
				fileOutputStream = new FileOutputStream(sourceFile);
				wb.write(fileOutputStream);
				String base64Str = Base64Utils.GetImageStr(sourceFile.getAbsolutePath());
				return new Result(true, base64Str);
			} catch (Exception e) {
				e.printStackTrace();
				return new Result(false,"转换失败，联系管理员",null);
			} finally{
				if(fileOutputStream!=null){
					fileOutputStream.close();
				}
				if(sourceFile.exists()){
					sourceFile.delete();
				}
			}
		}else{
			return new Result(false,"页面代码为空",null);
		}
		
	}
	
	
	/**给接收文件流用的**/
	private Result insertAndWrite(File sourceFile, File targetFile, String ewmJson) throws IOException{
		if(StringUtils.isNotBlank(ewmJson)){
			ExcelDTO excel = JsonUtils.jsonToJavaBean(ewmJson, ExcelDTO.class);
			FileOutputStream fileOut = null;
			if(!sourceFile.exists()){
				return new Result(Boolean.FALSE,"excel源文件不存在",null);
			}
			
			if(!targetFile.getParentFile().exists()){
				targetFile.getParentFile().mkdirs();
				if(targetFile.exists()){
					targetFile.delete();
				}
			}
			
			try{
				Workbook wb = PoiUtils.getWeebWork(sourceFile.getAbsolutePath());
				XSSFWorkbook wb2 = (XSSFWorkbook) wb;
				//开始修改excel
				for (SheetDTO sheet : excel.getSheetList()) {
					for (int s = 0; s < wb2.getNumberOfSheets(); s++) {  
						XSSFSheet _sheet = wb2.getSheetAt(s);  
						if(s+1 == sheet.getSheetNum()){//找到sheet
							System.out.println(_sheet.getSheetName());
							if(sheet.getCopy2()!=null && StringUtils.isNotBlank(sheet.getCopiedSheetName())){
								XSSFSheet newsheet = (XSSFSheet) wb.createSheet(sheet.getCopiedSheetName());
								PoiUtils.copySheet(wb2, _sheet, newsheet, true);
							}else if(sheet.getDoRemove()!=null && sheet.getDoRemove()){
								wb2.removeSheetAt(s);
							}else if(CollectionUtils.isNotEmpty(sheet.getRowList())){
								
								int rowCount = _sheet.getPhysicalNumberOfRows(); //获取总行数  
								int startRowShift = 0;//指定起始行的偏移量
								for (RowDTO row : sheet.getRowList()) {
									rowCount = _sheet.getPhysicalNumberOfRows();//重新获取总行数
									row.shiftStartRow(startRowShift);
									if(row.getStartRow() <= rowCount){
										
										switch (row.getType()) {//从某行开始向下插入多行
										case "insert":{
											if(row.isWriteFromStart()){
												PoiUtils.insertAndWriteRowFromStart(wb2, _sheet, row.getStartRow(), row.getMyRows());
												startRowShift += row.getMyRows().size()-1;
											}else{
												PoiUtils.insertAndWriteRow(wb2, _sheet, row.getStartRow(), row.getMyRows());
												startRowShift += row.getMyRows().size();
											}
											break;
										}
										case "write":{//覆盖某行
											PoiUtils.writeRow(wb2, _sheet, row.getStartRow(), row.getMyRows());
											break;
										}
										case "copy":{//复制行
											PoiUtils.copyRow(wb2, _sheet, row.getSourceRows(),row.getTargetStartRow()+startRowShift, sheet.getRowList());
											startRowShift += row.getSourceRows().length;
											break;
										}
										default:
											break;
										}
										
									}else{
										return new Result(Boolean.FALSE,"起始行超出该sheet总行数",null);
									}
								}
							}
							
							
							break;
						}
					}
				}
				
				fileOut = new FileOutputStream(targetFile.getAbsoluteFile());   
				wb.write(fileOut); 
				return  new Result(Boolean.TRUE);
				
			} catch (Exception e) {
				e.printStackTrace();
				return new Result(Boolean.FALSE,"接口报错",null);
			} finally {
				if(fileOut!=null){
					fileOut.close();
				}
			}	
			
		}else{
			return new Result(Boolean.FALSE,"缺少json参数",null);
		}
	}
	
	@Deprecated
	@RequestMapping(value = "/insertRow", method = RequestMethod.POST)
	@ResponseBody
	public Result insertRow(@RequestParam String ewmJson, HttpServletRequest request, HttpServletResponse response) throws IOException{
		if(StringUtils.isNotBlank(ewmJson)){
			ExcelDTO excel = JsonUtils.jsonToJavaBean(ewmJson, ExcelDTO.class);
			if(StringUtils.isNotBlank(excel.getSourcePath()) && StringUtils.isNotBlank(excel.getTargetPath())){
				File sourceFile = new File(excel.getSourcePath());
				File targetFile = new File(excel.getTargetPath());
				FileOutputStream fileOut = null;
				if(!sourceFile.exists()){
					return new Result(Boolean.FALSE,"excel源文件不存在",null);
				}
				
				if(!targetFile.getParentFile().exists()){
					targetFile.getParentFile().mkdirs();
					if(targetFile.exists()){
						targetFile.delete();
					}
				}
				
				try{
					Workbook wb = PoiUtils.getWeebWork(excel.getSourcePath());
					XSSFWorkbook wb2 = (XSSFWorkbook) wb;
					
					//开始修改excel
					for (SheetDTO sheet : excel.getSheetList()) {
						int sheetCount = wb2.getNumberOfSheets();
						for (int s = 0; s < sheetCount; s++) {  
							XSSFSheet _sheet = wb2.getSheetAt(s);  
							if(s+1 == sheet.getSheetNum()){//找到sheet
								int rowCount = _sheet.getPhysicalNumberOfRows(); //获取总行数  
								int startRowShift = 0;//指定起始行的偏移量
								for (RowDTO row : sheet.getRowList()) {
									row.shiftStartRow(startRowShift);
									if(row.getStartRow() <= rowCount){
										
										PoiUtils.insertRow(wb2, _sheet, row.getStartRow(), row.getInsertRow());
										
										startRowShift += row.getInsertRow();
									}else{
										return new Result(Boolean.FALSE,"起始行超出该sheet总行数",null);
									}
								}
								
								break;
							}
						}
					}
					
					fileOut = new FileOutputStream(excel.getTargetPath());   
					wb.write(fileOut); 
					return new Result(Boolean.TRUE);
					
				} catch (Exception e) {
					e.printStackTrace();
					return new Result(Boolean.FALSE,"接口报错",null);
				} finally {
					if(fileOut!=null){
						fileOut.close();
					}
				}
				
				
			}
			
		}else{
			return new Result(Boolean.FALSE,"缺少json参数",null);
		}
		
		
		return new Result(Boolean.FALSE,"缺少参数",null);
	}
	
	@Deprecated
	@RequestMapping(value = "/edit", method = RequestMethod.POST)
	@ResponseBody
	public Result insertAndWriteRow(@RequestParam String ewmJson, HttpServletRequest request, HttpServletResponse response) throws IOException{
		if(StringUtils.isNotBlank(ewmJson)){
			ExcelDTO excel = JsonUtils.jsonToJavaBean(ewmJson, ExcelDTO.class);
			if(StringUtils.isNotBlank(excel.getSourcePath()) && StringUtils.isNotBlank(excel.getTargetPath())){
				File sourceFile = new File(excel.getSourcePath());
				File targetFile = new File(excel.getTargetPath());
				FileOutputStream fileOut = null;
				if(!sourceFile.exists()){
					return new Result(Boolean.FALSE,"excel源文件不存在",null);
				}
				
				if(!targetFile.getParentFile().exists()){
					targetFile.getParentFile().mkdirs();
					if(targetFile.exists()){
						targetFile.delete();
					}
				}
				
				try{
					Workbook wb = PoiUtils.getWeebWork(excel.getSourcePath());
					XSSFWorkbook wb2 = (XSSFWorkbook) wb;
					
					//开始修改excel
					for (SheetDTO sheet : excel.getSheetList()) {
						int sheetCount = wb2.getNumberOfSheets();
						for (int s = 0; s < sheetCount; s++) {  
							XSSFSheet _sheet = wb2.getSheetAt(s);  
							if(s+1 == sheet.getSheetNum()){//找到sheet
								System.out.println(_sheet.getSheetName());
								if(sheet.getCopy2()!=null && StringUtils.isNotBlank(sheet.getCopiedSheetName())){
									XSSFSheet newsheet = (XSSFSheet) wb.createSheet(sheet.getCopiedSheetName());
									PoiUtils.copySheet(wb2, _sheet, newsheet, true);
								}else if(CollectionUtils.isNotEmpty(sheet.getRowList())){
									
									int rowCount = _sheet.getPhysicalNumberOfRows(); //获取总行数  
									int startRowShift = 0;//指定起始行的偏移量
									for (RowDTO row : sheet.getRowList()) {
										row.shiftStartRow(startRowShift);
										if(row.getStartRow() <= rowCount){
											
											switch (row.getType()) {//从某行开始向下插入多行
											case "insert":{
												if(row.isWriteFromStart()){
													PoiUtils.insertAndWriteRowFromStart(wb2, _sheet, row.getStartRow(), row.getMyRows());
													startRowShift += row.getMyRows().size()-1;
												}else{
													PoiUtils.insertAndWriteRow(wb2, _sheet, row.getStartRow(), row.getMyRows());
													startRowShift += row.getMyRows().size();
												}
												break;
											}
											case "write":{//覆盖某行
												PoiUtils.writeRow(wb2, _sheet, row.getStartRow(), row.getMyRows());
												break;
											}
											case "copy":{//复制行
												PoiUtils.copyRow(wb2, _sheet, row.getSourceRows(),row.getTargetStartRow()+startRowShift, sheet.getRowList());
												startRowShift += row.getSourceRows().length;
												break;
											}
											default:
												break;
											}
											
										}else{
											return new Result(Boolean.FALSE,"起始行超出该sheet总行数",null);
										}
									}
								}
								
								
								break;
							}
						}
					}
					
					fileOut = new FileOutputStream(excel.getTargetPath());   
					wb.write(fileOut); 
					return  new Result(Boolean.TRUE);
					
				} catch (Exception e) {
					e.printStackTrace();
					return new Result(Boolean.FALSE,"接口报错",null);
				} finally {
					if(fileOut!=null){
						fileOut.close();
					}
				}
			}
			
		}else{
			return new Result(Boolean.FALSE,"缺少json参数",null);
		}
		
		return new Result(Boolean.FALSE,"缺少参数",null);
	}
	
	
	@Deprecated
	@RequestMapping(value = "/writeRow", method = RequestMethod.POST)
	@ResponseBody
	public Result writeRow(@RequestParam String ewmJson, HttpServletRequest request, HttpServletResponse response) throws IOException{
		if(StringUtils.isNotBlank(ewmJson)){
			ExcelDTO excel = JsonUtils.jsonToJavaBean(ewmJson, ExcelDTO.class);
			if(StringUtils.isNotBlank(excel.getSourcePath()) && StringUtils.isNotBlank(excel.getTargetPath())){
				File sourceFile = new File(excel.getSourcePath());
				File targetFile = new File(excel.getTargetPath());
				FileOutputStream fileOut = null;
				if(!sourceFile.exists()){
					return new Result(Boolean.FALSE,"excel源文件不存在",null);
				}
				
				if(!targetFile.getParentFile().exists()){
					targetFile.getParentFile().mkdirs();
					if(targetFile.exists()){
						targetFile.delete();
					}
				}
				
				try{
					Workbook wb = PoiUtils.getWeebWork(excel.getSourcePath());
					XSSFWorkbook wb2 = (XSSFWorkbook) wb;
					
					//开始修改excel
					for (SheetDTO sheet : excel.getSheetList()) {
						int sheetCount = wb2.getNumberOfSheets();
						for (int s = 0; s < sheetCount; s++) {  
							XSSFSheet _sheet = wb2.getSheetAt(s);  
							if(s+1 == sheet.getSheetNum()){//找到sheet
								int rowCount = _sheet.getPhysicalNumberOfRows(); //获取总行数  
								int startRowShift = 0;//指定起始行的偏移量
								for (RowDTO row : sheet.getRowList()) {
									row.shiftStartRow(startRowShift);
									if(row.getStartRow() <= rowCount){
										
										PoiUtils.writeRow(wb2, _sheet, row.getStartRow(), row.getMyRows());
										
										startRowShift += row.getMyRows().size();
									}else{
										return new Result(Boolean.FALSE,"起始行超出该sheet总行数",null);
									}
								}
								
								break;
							}
						}
					}
					
					fileOut = new FileOutputStream(excel.getTargetPath());   
					wb.write(fileOut); 
					return new Result(Boolean.TRUE);
					
				} catch (Exception e) {
					e.printStackTrace();
					return new Result(Boolean.FALSE,"接口报错",null);
				} finally {
					if(fileOut!=null){
						fileOut.close();
					}
				}
			}
			
		}else{
			return new Result(Boolean.FALSE,"缺少json参数",null);
		}
		
		return new Result(Boolean.FALSE,"缺少参数",null);
	}
}
