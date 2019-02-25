package com.dvt.PoiService.commons.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.net.URL;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlrpc.client.XmlRpcClient;
import org.apache.xmlrpc.client.XmlRpcClientConfigImpl;
import org.junit.Test;

import com.dvt.PoiService.business.main.dto.ExcelDTO;
import com.dvt.PoiService.business.main.dto.RowDTO;
import com.dvt.PoiService.business.main.dto.SheetDTO;
import com.dvt.PoiService.commons.entity.Result;
import com.dvt.PoiService.commons.utils.CommonHelper;
import com.dvt.PoiService.commons.utils.HttpHelper;
import com.dvt.PoiService.commons.utils.JsonUtils;
import com.dvt.PoiService.commons.utils.POIConvertHtml2Excel;
import com.dvt.PoiService.commons.utils.POIReadExcelToHtml;
import com.dvt.PoiService.commons.utils.PoiUtils;
import com.dvt.PoiService.commons.utils.XmlUtils;
import com.google.common.collect.ImmutableList;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

public class TestMethod {
	
	public static Integer getCharNum(String str){
		String r = "0";
		for(int i=0;i<str.length();i++){
			if(str.charAt(i)>=48 && str.charAt(i)<=57){
				r+=str.charAt(i);
			}
		}
		return Integer.valueOf(r);
	}
	public static String getCharEn(String str){
		String r = "";
		for(int i=0;i<str.length();i++){
			if(str.charAt(i)>=65 && str.charAt(i)<=90){
				r+=str.charAt(i);
			}
		}
		return r;
	}
	
	
	@Test
	public void test33() throws IOException{
	}
	
	//@Test
    public void test22() throws IOException{
    	String readPath = "D://QMES//761043617//FileRecv//大客部3月份考勤.xlsx";
    	String writePath = "D://QMES//761043617//FileRecv//大客部3月份考勤.html";
    	int sheetNum = 0;
    	
    	File writeFile = new File(writePath);
        if(writeFile.exists()){
        	writeFile.delete();
        	writeFile.createNewFile();
        }
    	
    	File sourcefile = new File(readPath);
    	InputStream is = new FileInputStream(sourcefile);
        XSSFWorkbook wb = null;
        try {
        	wb = (XSSFWorkbook)WorkbookFactory.create(is);
        	XSSFSheet sheet = wb.getSheetAt(sheetNum);
        	
        	String html = POIReadExcelToHtml.readExcelToHtml(wb, sheet, true);
            System.out.println(html);
            
            html = "<html><head></head><body>" + html + "</body></html>";
            FileUtils.writeStringToFile(writeFile, XmlUtils.formatHtml(html));
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
           is.close();
        }
    	
    }
	
	//@Test
	public void testHttp(){
		Map<String, String> params = Maps.newHashMap();
		params.put("data", "{\"address\":\"test-xsxx\",\"params\": {},\"token\":\"C1A035457E5CF01B69A4270529F86638\"}");
		String s1;
		try {
			s1 = HttpHelper.startPost("http://222.197.182.5:7018/datacenter/rest/dataservicesearch/QueryServiceData", params);
			System.out.println(s1);
		} catch (IOException e) {
			e.printStackTrace();
		}
				
	}
	
	//@Test
	public void test(){
		File file = new File("D:/test.20170818");
		List<String> citys = ImmutableList.of("Beijing","Shanghai","Chongqing","Tianjin","Shijiazhuang","Baoding","Hangzhou","Shenyang","Kunming","Shenzhen","Guangzhou");
		List<String> months = ImmutableList.of("201701","201702","201703","201704","201705","201706","201707","201708","201709","201710","201711","201712");
		try {
			if(!file.exists()&&file.isFile()){
				file.createNewFile();
			}
			List<String> cache = Lists.newArrayList();
			while(file.length()<1048576L*1024){
				String city  = citys.get(new Random().nextInt(citys.size()));
				String month  = months.get(new Random().nextInt(months.size()));
				Long count =  (long)(Math.random() * 100000);
				String line = city + "|" + month + "|" + count;
				cache.add(line);
				if(cache.size()==100){
					FileUtils.writeLines(file, cache, Boolean.TRUE);
					cache.clear();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	//@Test
	public void test2() throws Exception{
		final String URL = "http://localhost:8069"; 
	    //final String DB = "hospital-mh";  
	    //final int USERID = 1;  
	    //final String PASS = "123456";  
	    List<String> emptyList = Lists.newArrayList();
		
	    XmlRpcClientConfigImpl config = new XmlRpcClientConfigImpl();  
		XmlRpcClient client = new XmlRpcClient();
		
	    config.setServerURL(new URL(String.format("%s/xmlrpc/2/common", URL)));  
	    client.setConfig(config);
		Object obj = client.execute(config, "authenticate", Arrays.asList(
        "hospital-mh", "admin", "123456", Maps.newHashMap()));
		 
        
        if(obj != null){
        	System.out.println(obj.toString());
        }
	}
	
//	@Test
//	public void test3() throws Exception{
//		try {
//			Workbook wb = PoiUtils.getWeebWork("D:/QMES/761043617/FileRecv/工作簿1(1).xlsx");
//			XSSFWorkbook wb2 = (XSSFWorkbook) wb;
//			int sheetCount = wb.getNumberOfSheets();  //Sheet的数量  
//			for (int s = 0; s < sheetCount; s++) {  
//				XSSFSheet sheet = wb2.getSheetAt(s);  
//				if("Sheet12".equals(sheet.getSheetName())){
//					//PoiUtils.insertRow(wb2, sheet, 17, 1);
//					
//					List<String> row1 = ImmutableList.of("1","2","3","4","5");
//					List<String> row2 = ImmutableList.of("6","7","8","9","10");
//					List<List<String>> myrow = Lists.newArrayList();
//					myrow.add(row1);
//					myrow.add(row2);
//					PoiUtils.writeRow(wb2, sheet, 17, myrow);
//				}
//				int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数  
//				System.out.println(rowCount);
//			}
//			FileOutputStream fileOut = new FileOutputStream("D:/QMES/761043617/FileRecv/工作簿1(1)_new.xlsx");   
//			wb.write(fileOut); 
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//		
//	}
	//@Test
	public void test4() throws Exception{
		try {
			
			List<RowDTO> dtoList = Lists.newArrayList();
			List<String> row1 = ImmutableList.of("1","2","3","4","5");
			List<String> row2 = ImmutableList.of("6","7","8","9","10");
			List<List<String>> myrow1 = Lists.newArrayList();
			myrow1.add(row1);
			myrow1.add(row2);
			dtoList.add(new RowDTO(4,  myrow1, "insert", true));
			
			List<List<String>> myrow2 = Lists.newArrayList();
			List<String> row3 = ImmutableList.of("11","12","13","14","None","15");
			myrow2.add(row3);
			dtoList.add(new RowDTO(5,  myrow2, "write"));
			
//			int[] aa = {4,5};
//			dtoList.add(new RowDTO(aa,6));
			
			List<SheetDTO> sheetList = Lists.newArrayList();
			//sheetList.add(new SheetDTO(3, dtoList));
			sheetList.add(new SheetDTO(3, 6, "新项目"));
			ExcelDTO excel = new ExcelDTO("D:/QMES/761043617/FileRecv/XXXX年研发费用辅助账-高新.xlsx", "D:/QMES/761043617/FileRecv/XXXX年研发费用辅助账-高新_new.xlsx", sheetList);
			String ewmJson = JsonUtils.JavaBeanToJson(excel);
			System.out.println(ewmJson);
			//
			
			if(StringUtils.isNotBlank(excel.getSourcePath()) && StringUtils.isNotBlank(excel.getTargetPath())){
				File sourceFile = new File(excel.getSourcePath());
				File targetFile = new File(excel.getTargetPath());
				FileOutputStream fileOut = null;
				if(!sourceFile.exists()){
					 new Result(Boolean.FALSE,"excel源文件不存在",null);
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
											new Result(Boolean.FALSE,"起始行超出该sheet总行数",null);
										}
									}
								}
								
								
								break;
							}
						}
					}
					
					fileOut = new FileOutputStream(excel.getTargetPath());   
					wb.write(fileOut); 
					 new Result(Boolean.TRUE);
					
				} catch (Exception e) {
					e.printStackTrace();
					 new Result(Boolean.FALSE,"接口报错",null);
				} finally {
					if(fileOut!=null){
						fileOut.close();
					}
				}
			}
			
			
			
			
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
}
