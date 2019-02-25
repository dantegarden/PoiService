package com.dvt.PoiService.commons.utils;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;

/**
 * 通用辅助类
 * 
 * @author AcoreHeng
 */
public final class CommonHelper {
	private final static String EMPTY = "";
	private final static String SEPARATOR_COMMA = ",";
	public final static String DF_DATE_TIME = "yyyy-MM-dd HH:mm:ss";
	public final static String DF_DATE_SHORT_TIME = "yyyy-MM-dd HH:mm";
	public final static String DF_DATE = "yyyy-MM-dd";
	public final static String DF_TIME = "HH:mm:ss";
	public final static String DF_DATE_SUB = "yyyyMMdd";
	public final static String DF_DATE_SUB_TIME = "yyyyMMddHHmmss";
	private static Map<String, DateFormat> DF_MAP;
	static {
		DF_MAP = new HashMap<String, DateFormat>();
		DF_MAP.put(DF_DATE, new SimpleDateFormat(DF_DATE));
		DF_MAP.put(DF_DATE_SHORT_TIME, new SimpleDateFormat(DF_DATE_SHORT_TIME));
		DF_MAP.put(DF_DATE_TIME, new SimpleDateFormat(DF_DATE_TIME));
	}
	
	public static Date Str2DateAutoFormat(String strValue) {
		if (StringUtils.isBlank(strValue) || strValue.contains("*")) {
			return null;
		}
		SimpleDateFormat clsFormat = null;
		if (strValue.length() > 19)
			strValue = strValue.substring(0, 19);
		if (strValue.length() == 19)
			clsFormat = new SimpleDateFormat(CommonHelper.DF_DATE_TIME);
		else if (strValue.length() == 21 && strValue.contains(".")){
			strValue = strValue.substring(0,19);
			clsFormat = new SimpleDateFormat(CommonHelper.DF_DATE_TIME);
		}
		else if (strValue.length() == 10)
			clsFormat = new SimpleDateFormat(CommonHelper.DF_DATE);
		else if (strValue.length() == 14 && !strValue.contains("-"))
			clsFormat = new SimpleDateFormat(CommonHelper.DF_DATE_SUB_TIME);
		else if (strValue.length() == 8)
			clsFormat = new SimpleDateFormat(CommonHelper.DF_DATE_SUB);
		
		ParsePosition pos = new ParsePosition(0);
		return clsFormat == null ? null : clsFormat.parse(strValue, pos);
	}
	
	
	public boolean isDate(String date){
      /**
       * 判断日期格式和范围
       */
      String rexp = "^((\\d{2}(([02468][048])|([13579][26]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])))))|(\\d{2}(([02468][1235679])|([13579][01345789]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|(1[0-9])|(2[0-8]))))))";
      Pattern pat = Pattern.compile(rexp);
      Matcher mat = pat.matcher(date);
      boolean dateType = mat.matches();
      return dateType;
    }
	
	// TODO String
	/**
	 * 判断字符串不能是null、""、"null"(忽略大小写)
	 * 
	 * @param str
	 * @return
	 */
	public static boolean isNotNull(String str) {
		return StringUtils.isNoneBlank(str) && !"null".equalsIgnoreCase(str);
	}

	/**
	 * Object转String
	 * 
	 * @param obj
	 * @return
	 */
	public static String toStr(Object obj) {
		if (obj == null) {
			return null;
		}
		return String.valueOf(obj);
	}

	// TODO Long
	/**
	 * Object转Long
	 * 
	 * @param obj
	 * @return
	 */
	public static Long toLong(Object obj) {
		if (obj == null) {
			return null;
		}
		if (obj instanceof BigDecimal) {
			return ((BigDecimal) obj).longValue();
		}
		if (obj instanceof Long) {
			return (Long) obj;
		}
		return Long.parseLong(String.valueOf(obj));
	}

	/**
	 * Object转Int
	 * 
	 * @param obj
	 * @return
	 */
	public static Integer toInt(Object obj) {
		if (obj == null) {
			return null;
		}
		if (obj instanceof BigDecimal) {
			return ((BigDecimal) obj).intValue();
		}
		if (obj instanceof String) {
			return Integer.parseInt(String.valueOf(obj));
		}
		return (Integer) obj;
	}

	/**
	 * Object转Int
	 * 
	 * @param obj
	 * @return
	 */
	public static Integer toInt(Object obj, int defaultVal) {
		Integer val = toInt(obj);
		return val != null ? val : 0;
	}

	/**
	 * Object转BigDecimal
	 * 
	 * @param obj
	 * @return
	 */
	public static BigDecimal toBigDecimal(Object obj) {
		if (obj == null) {
			return null;
		}
		if (obj instanceof String) {
			String str = toStr(obj);
			if (StringUtils.isBlank(str)) {
				return null;
			}
			return new BigDecimal(toStr(obj));
		}
		if (obj instanceof Integer) {
			return new BigDecimal(toStr(obj));
		}
		return (BigDecimal) obj;
	}
	/**
	 * @param obj
	 * @param defaultVal
	 * @return
	 */
	public static BigDecimal toBigDecimal(Object obj,BigDecimal defaultVal) {
		if (obj == null) {
			return defaultVal;
		}
		if (obj instanceof String) {
			String str = toStr(obj);
			if (StringUtils.isBlank(str)) {
				return defaultVal;
			}
			return new BigDecimal(toStr(obj));
		}
		if (obj instanceof Integer) {
			return new BigDecimal(toStr(obj));
		}
		return (BigDecimal) obj;
	}
	
	/**
	 * 数组转字符串
	 * */
	public static String LObject2String(Object[] object){
		String str = "[]";
		if(object!=null){
			str = "[";
			for (Object obj : object) {
				if (obj instanceof Object[]){
					str+= LObject2String((Object[])obj)+",";
				}else{
					str+= str.toString() +",";
				}
			}
			if(object.length>0){
				str = str.substring(0, str.length()-2);
			}
			str+="]";
			return str;
		}
		return str;
	}
	/**
	 * String转BigDecimal
	 * 
	 * @param str
	 * @return
	 */
	public static BigDecimal str2BigDecimal(String str) {
  		if (str == null||"".equals(str)) {
			return BigDecimal.ZERO;
		}
		if (str instanceof String) {
			return new BigDecimal(toStr(str));
		}
		return null;
	}

	/**
	 * Long转Str
	 * 
	 * @param str
	 *            如果为null
	 * @return 如果为null则返回""
	 */
	public static String long2StrIfEmpty(Long val) {
		if (val == null) {
			return EMPTY;
		}
		return String.valueOf(val);
	}
	
	/**
	 * @param bds
	 * @return
	 */
	public static BigDecimal sumOfBigDecimal(BigDecimal... bds){
		if(bds==null||bds.length==0){
			return null;
		}
		BigDecimal r_bd=BigDecimal.ZERO;
		if(bds.length>0){
			for (BigDecimal bd : bds) {
				if(bd!=null){
					r_bd=r_bd.add(bd);
				}
			}
		}
		return r_bd;
	}
	/**
	 * 求差
	 * @param bd
	 * @param bds
	 * @return
	 */
	public static BigDecimal diffOfBigDecimal(BigDecimal bd,BigDecimal... bds){
		if(bd==null||bds==null||bds.length==0){
			return bd;
		}
		BigDecimal r_bd=bd;
		if(bds.length>0){
			for (BigDecimal temp_bd : bds) {
				if(bd!=null){
					r_bd=r_bd.subtract(temp_bd);
				}
			}
		}
		return r_bd;
	}

	// TODO Date
	/**
	 * 获取系统当前时间
	 * 
	 * @return
	 */
	public static Date getNow() {
		return new Date();
	}
	/**
	 * 获取系统当前时间的字符串格式
	 * 
	 * @return
	 */
	public static String getNowStr(String format) {
		if (StringUtils.isBlank(format)) {
			return null;
		}
		DateFormat df = getDateFormat(format);
		return df.format(new Date());
	}
	
	/**
	 * 获取指定格式DateFormat
	 * 
	 * @param format
	 * @return
	 */
	public static DateFormat getDateFormat(String format) {
		if (StringUtils.isBlank(format)) {
			return null;
		}
		DateFormat df = DF_MAP.get(format);
		if (df == null) {
			df = new SimpleDateFormat(format);
		}
		return df;
	}
	

	/**
	 * 格式化日期为指定格式
	 * 
	 * @param date
	 * @param format
	 * @return
	 */
	public static Date formatDate(final Date date, String format) {
		if (date == null) {
			return null;
		}
		if (StringUtils.isBlank(format)) {
			return date;
		}
		Date temp_date = null;
		try {
			DateFormat df = getDateFormat(format);
			temp_date = df.parse(df.format(date));
		} catch (ParseException e) {
			// e.printStackTrace();
		}
		return temp_date;
	}

	/**
	 * 将日期转换为指定格式的字符串
	 * 
	 * @param date
	 * @param format
	 * @return
	 */
	public static String date2Str(Date date, String format) {
		if (date == null || StringUtils.isBlank(format)) {
			return null;
		}
		DateFormat df = getDateFormat(format);
		return df.format(date);
	}

	/**
	 * 将日期字符串按照指定格式转换为date
	 * 
	 * @param dateStr
	 * @param format
	 * @return
	 */
	public static Date str2Date(String dateStr, String format) {
		if (StringUtils.isBlank(dateStr) || StringUtils.isBlank(format)) {
			return null;
		}
		try {
			DateFormat df = getDateFormat(format);
			return df.parse(dateStr);
		} catch (ParseException e) {
			// e.printStackTrace();
		}
		return null;
	}
	
	/**
	 * 将逗号隔开的字符串转为Set集合
	 * @param str
	 * @return
	 */
	public static Set<String> str2Set(String str){
		return new HashSet<String>(Arrays.asList(StringUtils.split(str,SEPARATOR_COMMA)));
	}
	/**
	 * @param str
	 * @return
	 */
	public static Set<Long> str2SetOfLong(String str){
		String[] strs=StringUtils.split(str,SEPARATOR_COMMA);
		Set<Long> set=new HashSet<Long>();
		if(strs!=null&&strs.length>0){
			for (String s : strs) {
				if(StringUtils.isNumeric(s)){
					set.add(toLong(s));
				}
			}
		}
		return set;
	}
	public static String set2Str(Set<String> set){
		return StringUtils.join(set.toArray(),SEPARATOR_COMMA);
	}
	/**
	 * 将逗号隔开的字符串转为List集合
	 * @param str
	 * @return
	 */
	public static List<String> str2List(String str){
		return Arrays.asList(StringUtils.split(str,SEPARATOR_COMMA));
	}

	/**
	 * Object转Boolean
	 * 
	 * @param obj
	 * @return
	 */
	public static Boolean toBoolean(Object obj) {
		if (obj == null) {
			return null;
		}
		if (obj instanceof Boolean) {
			return (Boolean) obj;
		}
		return null;
	}
	
	public static Date toDate(Object obj) {
		if (obj == null) {
			return null;
		}
		if (obj instanceof Date) {
			return (Date) obj;
		}
		if (obj instanceof Timestamp) {
			return new Date(((Timestamp) obj).getTime());
		}
		return null;
	}

	// TODO　test
	public static void main(String[] args) {
		System.out.println(date2Str(getNow(), "yyyy-MM-dd HH:mm:ss"));
		System.out.println(str2BigDecimal("3"));
		System.out.println(str2Set("2342342,23稍等42342,热特委托我"));
	}

	/**
     * 文件重命名
     */
    public static String renameFileName(String fileName) {
        String formatDate = new SimpleDateFormat("yyMMddHHmmss").format(new Date()); // 当前时间字符串
        int random = new Random().nextInt(10000);
        String extension = fileName.substring(fileName.lastIndexOf(".")); // 文件后缀
 
        return formatDate + random + extension;
    }
    
    /** 
     * 校验字符串是否是数值(包含小数与负数) 
     * 示例： 
     * false : . 1. 1sr -  12. -12. 
     * true: -12 -12.0 -12.056 12 12.0 12.056 
     * 
     * @param str 需要校验的字符串 
     * @return false :不是数值 true：是数值 6
     */  
    public static Boolean checkNumber(String str) {  
        String regex = "[-]?[0-9]+(\\.{1}[0-9]+)?";  
        if (str == null || !str.matches(regex)) {  
            return false;  
        }  
        return true;  
    }  
	
}
