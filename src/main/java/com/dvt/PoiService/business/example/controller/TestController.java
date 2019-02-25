package com.dvt.PoiService.business.example.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.dvt.PoiService.business.example.service.TestService;
import com.dvt.PoiService.business.main.dto.FormDTO;
import com.dvt.PoiService.commons.GlobalConstants;
import com.dvt.PoiService.commons.utils.FileUtils;
import com.dvt.PoiService.commons.utils.XmlRpcUtils;
import com.dvt.PoiService.commons.utils.XmlUtils;
import com.dvt.PoiService.commons.vo.DownloadRecord;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;



@Controller
@RequestMapping("/test")
public class TestController {
	private static final Logger logger = LoggerFactory.getLogger(TestController.class);
	
	@Autowired
	private TestService testService;
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	@RequestMapping
	public String init(HttpServletRequest request) {
		System.out.println("test initting");
		//项目IP，数据库，模型名,action_id,
		
		Integer uid = 1;
		Integer action_id = 451; 
		String odooUrl = "http://101.200.124.206:4713";
		String db = "hospital-saas";
		String pwd = "123456";
		String model = "maintenance.equipment";
		
		String url = String.format("%s/xmlrpc/2/object", odooUrl);
		
		Map context = Maps.newHashMap();
		context.put("lang", "zh_CN");
		context.put("tz", "Asia/Hong_Kong");
		context.put("uid", uid);
		context.put("search_default_no_share", 1);
		context.put("params", ImmutableMap.of("action",action_id));
		
		Map options = Maps.newHashMap();
		options.put("action_id", action_id);
		options.put("load_fields", true);
		options.put("toolbar", true);
		options.put("load_filters", true);
		
		List views = Lists.newArrayList(
										Lists.newArrayList(884,"list"),
										Lists.newArrayList(883,"form"),
										Lists.newArrayList(933,"search")
										);
		
		Map loadViews = (HashMap) XmlRpcUtils.getXMLRPC(url, "execute_kw", Arrays.asList(db, uid, pwd, model,"load_views",
				Lists.newArrayList(),new ImmutableMap.Builder<String, Object>()
									 .put("context", context)
									 .put("options", options)
									 .put("views", views).build()
		));
		
		Map fieldsViews = (HashMap)loadViews.get("fields_views");
		Map fields = (HashMap)loadViews.get("fields");  //模型的字段
		Object[] filters = (Object[]) loadViews.get("filters"); 
		
		//打印模型的字段
		for (Object key : fields.keySet()) {
			System.out.println(key+":");
			Map field = (HashMap)(fields.get(key));
			for (Object _key : field.keySet()) {
				System.out.println("   -- "+_key+":" + field.get(_key));
			}
		}
		
		//视图
		Map list = null;
		Map search = null;
		Map form = null;
		if(fieldsViews.containsKey("list")){
			list = (HashMap)fieldsViews.get("list");
		}
		if(fieldsViews.containsKey("search")){
			search = (HashMap)fieldsViews.get("search");
		}
		if(fieldsViews.containsKey("form")){
			form = (HashMap)fieldsViews.get("form");
		}
		
		
		String timestamp = "templates." + new Date().getTime();
		String folderPath = request.getSession().getServletContext().getRealPath("")+"template\\" + timestamp;
		File templateFolder = new File(folderPath);
		if(!templateFolder.exists()){
			templateFolder.mkdir();
			XmlUtils.toXMLFile(XmlUtils.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"+ list.get("arch").toString()), folderPath+"\\tree.xml", null);
			XmlUtils.toXMLFile(XmlUtils.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"+ form.get("arch").toString()), folderPath+"\\form.xml", null);
			XmlUtils.toXMLFile(XmlUtils.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"+ search.get("arch").toString()), folderPath+"\\search.xml", null);
			List<Node> VisibleFields = XmlUtils.XpathQuery(XmlUtils.read(folderPath+"\\tree.xml"),"//tree/field[not(@invisible)]");
			for (Node node : VisibleFields) {
				Element e = (Element)node;
				System.out.println(e.attributeValue("name"));
			}
		}
		
		return GlobalConstants.PAGE_AUTOCODE;
	}
	
	@RequestMapping("/download")
	@ResponseBody
	public DownloadRecord download(@ModelAttribute FormDTO dto, HttpServletRequest request, HttpServletResponse response) throws Exception {
		File f = new File(request.getSession().getServletContext().getRealPath("")+"template\\"+"templates.1501654071274"+"\\tree.xml");
		
		//声明本次下载状态的记录对象
		DownloadRecord downloadRecord = new DownloadRecord(f.getName(), f.getAbsolutePath(), request);
		//设置响应头和客户端保存文件名
		response.setCharacterEncoding("utf-8");
		response.setContentType("multipart/form-data");
		response.setHeader("Content-Disposition", "attachment;fileName=" + f.getName());
		//用于记录以完成的下载的数据量，单位是byte
		long downloadedLength = 0l;
		
		InputStream is = null;
		OutputStream os = null;
		
		try {
			//打开本地文件流
	        is = new FileInputStream(f.getAbsolutePath());
	        //激活下载操作
	        os = response.getOutputStream();
	        //循环写入输出流
	        byte[] b = new byte[2048];
	        int length;
	        while ((length = is.read(b)) > 0) {
	            os.write(b, 0, length);
	            downloadedLength += b.length;
	        }
	        
	        downloadRecord.setStatus(DownloadRecord.STATUS_SUCCESS);
	        downloadRecord.setEndTime(new Timestamp(System.currentTimeMillis()));
	        downloadRecord.setLength(downloadedLength);
	        
		} catch (Exception e) {
			downloadRecord.setStatus(DownloadRecord.STATUS_ERROR);
	        throw e;
		} finally {
			// 这里主要关闭。
			try {
				if(os!=null){
					os.close();
				}
				if(is!=null){
					is.close();
				}
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
		return downloadRecord;
	}
}
