package com.dvt.PoiService.business.main.controller;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.dom4j.Element;
import org.dom4j.Node;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;



import org.springframework.web.multipart.MultipartFile;

import com.dvt.PoiService.business.main.dto.FormDTO;
import com.dvt.PoiService.business.main.dto.UploadDTO;
import com.dvt.PoiService.commons.GlobalConstants;
import com.dvt.PoiService.commons.utils.CommonHelper;
import com.dvt.PoiService.commons.utils.FileUtils;
import com.dvt.PoiService.commons.utils.XmlRpcUtils;
import com.dvt.PoiService.commons.utils.XmlUtils;
import com.dvt.PoiService.commons.vo.DownloadRecord;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;


@Controller
@RequestMapping("/autocode")
public class AutoCodeController {
	private static final Logger logger = LoggerFactory.getLogger(AutoCodeController.class);
	
	private static final String LINE_SEP = System.getProperty("line.separator");
	private static final String PATH_SEP = File.separator;
	
	private static boolean hasInitTemplate = Boolean.FALSE;
	private static List<String> field_char_html;
	private static List<String> field_selection_html;
	private static List<String> field_many2one_html;
	private static List<String> field_date_html;
	private static List<String> form_html;
	
	private void initTemplate(HttpServletRequest request,String templatePath){
		try {
			field_char_html = FileUtils.readLines(new File(templatePath+PATH_SEP+"field_char.html"));
			field_selection_html = FileUtils.readLines(new File(templatePath+PATH_SEP+"field_selection.html"));
			field_many2one_html = FileUtils.readLines(new File(templatePath+PATH_SEP+"field_many2one.html"));
			field_date_html = FileUtils.readLines(new File(templatePath+PATH_SEP+"field_date.html"));
			form_html = FileUtils.readLines(new File(templatePath+PATH_SEP+"form.html"));
			hasInitTemplate = Boolean.TRUE;
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private String outputXml(HttpServletRequest request, Map fieldsViews){
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
		String folderPath = request.getSession().getServletContext().getRealPath("")+"template" + PATH_SEP + timestamp;
		File templateFolder = new File(folderPath);
		if(!templateFolder.exists()){
			templateFolder.mkdir();
			if(list!=null)
				XmlUtils.toXMLFile(XmlUtils.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + LINE_SEP + list.get("arch").toString()), folderPath+PATH_SEP+"tree.xml", null);
			if(form!=null)
				XmlUtils.toXMLFile(XmlUtils.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + LINE_SEP + form.get("arch").toString()), folderPath+PATH_SEP+"form.xml", null);
			if(search!=null)
				XmlUtils.toXMLFile(XmlUtils.parseText("<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + LINE_SEP + search.get("arch").toString()), folderPath+PATH_SEP+"search.xml", null);
		}
		return folderPath;
	}
	
	/**批量替换占位符*/
	private StringBuffer fillField(StringBuffer targetStr, List<Node> nodeFields, Map fields){
		for (Node node : nodeFields) {
			Element e = (Element)node;
			Map modelField = (HashMap)fields.get(e.attributeValue("name"));//某个模型字段
			String fieldName = (String)modelField.get("string");
			String fieldType = (String)modelField.get("type");
			if("char".equals(fieldType)||"integer".equals(fieldType)||"float".equals(fieldType)){//
				for (String line : field_char_html) {
					String _line = new String(line);
					_line = _line.replace("${field.string}", fieldName);
					_line = _line.replace("${field}",e.attributeValue("name"));
					targetStr.append(_line);
				}	
			}else if("selection".equals(fieldType)){
				for (String line : field_selection_html) {
					String _line = new String(line);
					_line = _line.replace("${field.string}", fieldName);
					_line = _line.replace("${field}",e.attributeValue("name"));
					targetStr.append(_line);
				}
			}else if("many2one".equals(fieldType)){
				for (String line : field_many2one_html) {
					String _line = new String(line);
					_line = _line.replace("${field.string}", fieldName);
					_line = _line.replace("${field}",e.attributeValue("name"));
					_line = _line.replace("${field.relation}", (String)modelField.get("relation"));
					_line = _line.replace("${field.domain}", CommonHelper.LObject2String((Object[])modelField.get("domain")));
					targetStr.append(_line);
				}
			}else if("date".equals(fieldType)||"datetime".equals(fieldType)){
				for (String line : field_date_html) {
					String _line = new String(line);
					_line = _line.replace("${field.string}", fieldName);
					_line = _line.replace("${field}",e.attributeValue("name"));
					_line = _line.replace("${format}", StringUtils.equals("date", fieldType)?CommonHelper.DF_DATE:CommonHelper.DF_DATE_TIME);
					targetStr.append(_line);
				}
			}
		}
		return targetStr;
	}
	
	@RequestMapping
	public String init(HttpServletRequest request) {
		String templatePath = request.getSession().getServletContext().getRealPath("")+"template";
		if(!hasInitTemplate){
			initTemplate(request, templatePath); //初始化模板
		}
		
		return GlobalConstants.PAGE_AUTOCODE;
	}
	
	@RequestMapping("/download")
	public void download(@ModelAttribute FormDTO dto, HttpServletRequest request, HttpServletResponse response){
		
		//——————————————————————————————准备参数————————————————————————————
		Integer uid = 1;
		Integer action_id = dto.getAction_id(); 
		String odooUrl = dto.getOdooUrl();
		String db = dto.getDb();
		String pwd = dto.getPwd();
		String model = dto.getModelName();
		
		String url = String.format("%s/xmlrpc/2/object", odooUrl);
		
		Map context = ImmutableMap.of(
			"lang", "zh_CN",
			"tz", "Asia/Hong_Kong",
			"uid", uid,
			"search_default_no_share", 1,
			"params", ImmutableMap.of("action",action_id)
		);
		
		Map options = ImmutableMap.of(
			"action_id", action_id,
			"load_fields", true,
			"toolbar", true,
			"load_filters", true
		);
		
		List views = Lists.newArrayList();
		if(dto.getForm_id()!=null)
			views.add(Lists.newArrayList(dto.getForm_id(),"form"));
		if(dto.getTree_id()!=null)
			views.add(Lists.newArrayList(dto.getTree_id(),"list"));
		if(dto.getSearch_id()!=null)
			views.add(Lists.newArrayList(dto.getSearch_id(),"search"));
		
		Map loadViews = (HashMap) XmlRpcUtils.getXMLRPC(url, "execute_kw", Arrays.asList(db, uid, pwd, model,"load_views",
				Lists.newArrayList(),new ImmutableMap.Builder<String, Object>()
									 .put("context", context)
									 .put("options", options)
									 .put("views", views).build()
		));
		
		Map fieldsViews = (HashMap)loadViews.get("fields_views");
		Map fields = (HashMap)loadViews.get("fields");  //模型的字段
		Object[] filters = (Object[]) loadViews.get("filters"); 
		
		String xmlPath = outputXml(request, fieldsViews);//落地模型视图的xml
		
		if(dto.getForm_id()!=null){
			String formFilePath = doForm(xmlPath, fieldsViews, fields);//填form 返回绝对路径
		
			//下载
			File f = new File(formFilePath);
			FileUtils.downloadFile(f, request, response);
		}
		
		
	}
	
	/**生成form文件*/
	private String doForm(String xmlpath, Map fieldsViews, Map fields) {
		Map form = null;
		String formXml = xmlpath + PATH_SEP + "form.xml";
		if(fieldsViews.containsKey("form")){
			form = (HashMap)fieldsViews.get("form");
			StringBuffer content = new StringBuffer("");

			//------------------表单头
			content.append("<!-- form -->");
			List<Node> titleFields = XmlUtils.XpathQuery(XmlUtils.read(formXml),"/form/sheet/div[@class='oe_title']//field[@name]");
			fillField(content, titleFields, fields);
			
			//------------------主表单
			List<Node> groupFields = XmlUtils.XpathQuery(XmlUtils.read(formXml),"/form/sheet/group[@string]");
			if(CollectionUtils.isNotEmpty(groupFields)){
				for (Node group : groupFields) {
					Element _group = (Element)group;
					String groupName = _group.attributeValue("string");
					content.append("<!-- "+groupName+" -->");
					content.append("<section class='mt10'>");
					List<Node> inGroupFields = _group.selectNodes("/form/sheet/group[@string='"+groupName+"']//field[@name]");
					fillField(content, inGroupFields, fields);
					content.append("</section>");
				}
			}
			
			//------------------TAB页
			List<Node> tabFields = XmlUtils.XpathQuery(XmlUtils.read(formXml),"/form/sheet/notebook/page[@string]");
			if(CollectionUtils.isNotEmpty(tabFields)){
				for (Node page : tabFields) {
					Element _page = (Element)page;
					String pageName = _page.attributeValue("string");
					content.append("<!-- "+pageName+" -->");
					content.append("<section class='mt10'>");
					List<Node> inGroupFields = _page.selectNodes("/form/sheet/notebook/page[@string='"+pageName+"']//field[@name]");
					fillField(content, inGroupFields, fields);
					content.append("</section>");
				}
			}
			
			
			//------------------生成文件
			String formFileName = "form_"+form.get("name")+".html";
			String result = "";
			try {
				File resultFile = new File(xmlpath + PATH_SEP + formFileName);
				if(!resultFile.exists())resultFile.createNewFile();
				for (String line : form_html) {
					String _line = new String(line);
					_line = _line.replace("${model.name}", (String)form.get("model"));
					_line = _line.replace("${fields}", content.toString());
					result += _line;
				}
				XmlUtils.parseText(result);//格式化HTML
				FileUtils.writeStringToFile(resultFile, XmlUtils.formatHtml(result));
			}catch(Exception e){
				e.printStackTrace();
			}
			//XmlUtils.toXMLFile(XmlUtils.parseText(result), xmlpath + PATH_SEP + formFileName, "UTF-8");
			return xmlpath + PATH_SEP + formFileName;
		}else{
			return null;
		}
		
	}
	
	
	@RequestMapping("/upload")
	public void upload(@ModelAttribute UploadDTO dto, HttpServletRequest request, HttpServletResponse response){
		String allowSuffix = "jpg,png,gif,jpeg";//允许文件格式
		
		String path = request.getContextPath();
		String basePath = request.getScheme() + "://" + request.getServerName() + ":" + 
	            request.getServerPort() + path;
		String realpath = request.getSession().getServletContext().getRealPath("");
		List<String> picPaths = Lists.newArrayList();
		try {
			String suffix = dto.getMultiFile().getOriginalFilename().substring(dto.getMultiFile().getOriginalFilename().lastIndexOf(".")+1);
			int length = allowSuffix.indexOf(suffix);
			if(length == -1){
				throw new Exception("请上传允许格式的文件");
			}
			File destFile = new File(realpath+"upload");
			if(!destFile.exists()){
				destFile.mkdirs();
			}
			System.out.println(dto.getMultiFile().getOriginalFilename());
			SimpleDateFormat fmt = new SimpleDateFormat("yyyyMMddHHmmssSSS");
			String fileNameNew = fmt.format(new Date())+"."+suffix;
			File f = new File(destFile.getAbsoluteFile()+"\\"+fileNameNew);
			dto.getMultiFile().transferTo(f);
			f.createNewFile();
			System.out.println(basePath+"upload//"+fileNameNew);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
