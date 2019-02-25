package com.dvt.PoiService.commons.utils;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIReadExcelToHtml {
	
	private static final String BLANK = "";//&nbsp;
	
	public static String getExcelInfo(Workbook wb, Sheet sheet, boolean isWithStyle){
        
        StringBuffer sb = new StringBuffer();
        if(sheet==null){
        	sheet = wb.getSheetAt(0);//获取第一个Sheet的内容
        }
        int lastRowNum = sheet.getLastRowNum();
        Map<String, String> map[] = getRowSpanColSpanMap(sheet);
        sb.append("<table id='poi_excel_to_html_table' style='border-collapse:collapse;' width='100%'>");
        Row row = null;        //兼容
        Cell cell = null;    //兼容
        
        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
            row = sheet.getRow(rowNum);
            if (row == null) {
                sb.append("<tr><td > "+ BLANK +"</td></tr>");
                continue;
            }
            sb.append("<tr>");
            int lastColNum = row.getLastCellNum();
            for (int colNum = 0; colNum < lastColNum; colNum++) {
                cell = row.getCell(colNum);
                if (cell == null) {    //特殊情况 空白的单元格会返回null
                    sb.append("<td>"+ BLANK +"</td>");
                    continue;
                }

                String stringValue = getCellValue(wb, cell);
                if (map[0].containsKey(rowNum + "," + colNum)) {
                    String pointString = map[0].get(rowNum + "," + colNum);
                    map[0].remove(rowNum + "," + colNum);
                    int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                    int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                    int rowSpan = bottomeRow - rowNum + 1;
                    int colSpan = bottomeCol - colNum + 1;
                    sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' ");
                } else if (map[1].containsKey(rowNum + "," + colNum)) {
                    map[1].remove(rowNum + "," + colNum);
                    continue;
                } else {
                    sb.append("<td ");
                }
                
                //判断是否需要样式
                if(isWithStyle){
                    dealExcelStyle(wb, sheet, cell, sb);//处理单元格样式
                }
                
                sb.append(">");
                if (stringValue == null || "".equals(stringValue.trim())) {
                    sb.append(" "+ BLANK +" ");
                } else {
                    // 将ascii码为160的空格转换为html下的空格（"+ BLANK +"）
                    sb.append(stringValue.replace(String.valueOf((char) 160),""+ BLANK +""));
                }
                sb.append("</td>");
            }
            sb.append("</tr>");
        }

        sb.append("</table>");
        return sb.toString();
    }

	private static Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {
	
	    Map<String, String> map0 = new HashMap<String, String>();
	    Map<String, String> map1 = new HashMap<String, String>();
	    int mergedNum = sheet.getNumMergedRegions();
	    CellRangeAddress range = null;
	    for (int i = 0; i < mergedNum; i++) {
	        range = sheet.getMergedRegion(i);
	        int topRow = range.getFirstRow();
	        int topCol = range.getFirstColumn();
	        int bottomRow = range.getLastRow();
	        int bottomCol = range.getLastColumn();
	        map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
	        // System.out.println(topRow + "," + topCol + "," + bottomRow + "," + bottomCol);
	        int tempRow = topRow;
	        while (tempRow <= bottomRow) {
	            int tempCol = topCol;
	            while (tempCol <= bottomCol) {
	                map1.put(tempRow + "," + tempCol, "");
	                tempCol++;
	            }
	            tempRow++;
	        }
	        map1.remove(topRow + "," + topCol);
	    }
	    Map[] map = { map0, map1 };
	    return map;
	}
	
	/**
     * 获取表格单元格Cell内容
     * @param cell
     * @return
     */
    private static String getCellValue(Workbook wb, Cell cell) {

        String result = new String();  
        switch (cell.getCellType()) {  
        case Cell.CELL_TYPE_NUMERIC:// 数字类型  
            if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式  
                SimpleDateFormat sdf = null;  
                if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {  
                    sdf = new SimpleDateFormat("HH:mm");  
                } else {// 日期  
                    sdf = new SimpleDateFormat("yyyy-MM-dd");  
                }  
                Date date = cell.getDateCellValue();  
                result = sdf.format(date);  
            } else if (cell.getCellStyle().getDataFormat() == 58) {  
                // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)  
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");  
                double value = cell.getNumericCellValue();  
                Date date = org.apache.poi.ss.usermodel.DateUtil  
                        .getJavaDate(value);  
                result = sdf.format(date);  
            } else {  
                double value = cell.getNumericCellValue();  
                CellStyle style = cell.getCellStyle();  
                DecimalFormat format = new DecimalFormat();  
                String temp = style.getDataFormatString();  
                // 单元格设置成常规  
                if (temp.equals("General")) {  
                    format.applyPattern("#");  
                }  
                result = format.format(value);  
            }  
            break;  
        case Cell.CELL_TYPE_FORMULA: //公式
        	result = "";
        	if(StringUtils.isNotBlank(cell.getCellFormula())){
        		XSSFFormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
            	CellValue c = formulaEvaluator.evaluate(cell);//getStringValue();
            	switch (c.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					result = c.getNumberValue() + "";
					break;
				case Cell.CELL_TYPE_STRING:
					result = c.getStringValue();
					break;
				default:
					break;
				}
        	}
        	break;
        case Cell.CELL_TYPE_STRING:// String类型  
            result = cell.getRichStringCellValue().toString().trim();  
            break; 
        case Cell.CELL_TYPE_BLANK:  
            result = "";  
            break; 
        default:  
            result = "";  
            break;  
        }  
        return result;  
    }
    
    /**
     * 处理表格样式
     * @param wb
     * @param sheet
     * @param cell
     * @param sb
     */
    private static void dealExcelStyle(Workbook wb,Sheet sheet,Cell cell,StringBuffer sb){
        
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            HorizontalAlignment alignment = cellStyle.getAlignmentEnum();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");//单元格内容的水平对齐方式
            VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignmentEnum();
            sb.append("valign='"+ convertVerticalAlignToHtml(verticalAlignment)+ "' ");//单元格中内容的垂直排列方式
            
            sb.append("style='");
            //加粗斜体
            XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont(); 
            boolean bold = xf.getBold();
            String boldWeight = (bold)?"bold":"normal";
            sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
            boolean italic = xf.getItalic();
            String fontItalic = (italic)?"italic":"normal";
            sb.append("font-style:" + fontItalic + ";");//斜体
            //字体颜色
            XSSFColor xc = xf.getXSSFColor();
            if (xc != null && !"".equals(xc)) {
                sb.append("color:#" + xc.getARGBHex().substring(2) + ";"); 
            }
            
            //背景颜色
            XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
            if (bgColor != null && !"".equals(bgColor)) {
                sb.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";"); 
            }
            //边框
            sb.append("border-left:"+ convertBorderStyleToHtml(cellStyle.getBorderLeft(),((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor(),cellStyle.getBorderLeftEnum()) +";");
            sb.append("border-top:"+ convertBorderStyleToHtml(cellStyle.getBorderTop(),((XSSFCellStyle) cellStyle).getTopBorderXSSFColor(),cellStyle.getBorderTopEnum()) +";");
            sb.append("border-right:"+ convertBorderStyleToHtml(cellStyle.getBorderRight(),((XSSFCellStyle) cellStyle).getRightBorderXSSFColor(),cellStyle.getBorderRightEnum()) +";");//
            sb.append("border-bottom:"+ convertBorderStyleToHtml(cellStyle.getBorderBottom(),((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor(),cellStyle.getBorderBottomEnum()) +";");//
            sb.append("' ");
        }
    }
    
    /**
     * 单元格内容的水平对齐方式
     * @param alignment
     * @return
     */
    private static String convertAlignToHtml(HorizontalAlignment alignment) {
        String align = "left";
        switch (alignment) {
        case LEFT:
            align = "left";
            break;
        case CENTER:
            align = "center";
            break;
        case RIGHT:
            align = "right";
            break;
        default:
            break;
        }
        return align;
    }
    
    private static String convertBorderStyleToHtml(short borderWidth, XSSFColor borderColor, BorderStyle borderStyle){
    	String bs = "";
    	switch(borderStyle){
    	case NONE:
    		bs = "none";
    		break;
    	case THIN:
    		bs = "solid";
    		break;
    	case MEDIUM:
    		bs = "solid";
    		break;
    	case DASHED:
    		bs = "dashed";
    		break;
    	case DOTTED:
    		bs = "dotted";
    		break;
    	case THICK:
    		bs = "thick";
    		break;
    	case DOUBLE:
    		bs = "double";
    		break;
    	case HAIR:
    		bs = "dotted";
    		break;
    	case MEDIUM_DASHED:
    		bs = "dashed";
    		break;
    	case DASH_DOT:
    		bs = "dotted";
    		break;
    	case MEDIUM_DASH_DOT:
    		bs = "dotted";
    		break;
    	case DASH_DOT_DOT:
    		bs = "dotted";
    		break;
    	case MEDIUM_DASH_DOT_DOT:
    		bs = "dotted";
    		break;
    	case SLANTED_DASH_DOT:
    		bs = "dotted";
    		break;
    	default:
            break;
        }
    	
    	//颜色
    	if (borderColor != null && !"".equals(borderColor)) {
            String borderColorStr = borderColor.getARGBHex();//t.getARGBHex();
            borderColorStr=borderColorStr==null|| borderColorStr.length()<1?" #000000":" #"+borderColorStr.substring(2);
            bs += borderColorStr;
        }
    	if(borderWidth>0){
    		bs += " 1px";
    	}
        return bs;
    }
    
    /**
     * 单元格中内容的垂直排列方式
     * @param verticalAlignment
     * @return
     */
    private static String convertVerticalAlignToHtml(VerticalAlignment verticalAlignment) {

        String valign = "middle";
        switch (verticalAlignment) {
        case BOTTOM:
            valign = "bottom";
            break;
        case CENTER:
            valign = "center";
            break;
        case TOP:
            valign = "top";
            break;
        default:
            break;
        }
        return valign;
    }
    
    private static String convertToStardColor(HSSFColor hc) {

        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }

        return sb.toString();
    }
    
    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }
    
    static String[] bordesr={"border-top:","border-right:","border-bottom:","border-left:"};
    static String[] borderStyles={"solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid","solid","solid","solid","solid"};

    private static  String getBorderStyle(HSSFPalette palette, int b, short s, short t){
         
        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
        String borderColorStr = convertToStardColor( palette.getColor(t));
        borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr;
        return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
        
    }
    
    private static  String getBorderStyle(int b,short s, XSSFColor xc){
         
         if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
         if (xc != null && !"".equals(xc)) {
             String borderColorStr = xc.getARGBHex();//t.getARGBHex();
             borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr.substring(2);
             return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
         }
         
         return "";
    }
    
	public static String readExcelToHtml(XSSFWorkbook xwb, XSSFSheet sheet, boolean isWithStyle){
        return POIReadExcelToHtml.getExcelInfo(xwb,sheet,isWithStyle);
	}
}
