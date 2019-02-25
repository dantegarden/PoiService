package com.dvt.PoiService.commons.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilderFactory;

import net.sf.jxls.util.Util;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import com.dvt.PoiService.business.main.dto.RowDTO;
import com.google.common.collect.Lists;

public class PoiUtils {
	private final static String excel2003L =".xls";    //2003- 版本的excel  
    private final static String excel2007U =".xlsx";   //2007+ 版本的excel  
    
    /** 
     *  
     * @Title: getWeebWork 
     * @Description: TODO(根据传入的文件名获取工作簿对象(Workbook)) 
     * @param filename 
     * @return 
     * @throws IOException 
     * @throws InvalidFormatException 
     * @throws EncryptedDocumentException 
     */  
    public static Workbook getWeebWork(String filename) throws IOException, EncryptedDocumentException, InvalidFormatException {  
    	Workbook workbook = null;  
        if (null != filename) {  
            String fileType = filename.substring(filename.lastIndexOf("."),  
                    filename.length());  
            FileInputStream fileStream = new FileInputStream(new File(filename));  
            if (excel2003L.equals(fileType.trim().toLowerCase())) {  
                workbook = new HSSFWorkbook(fileStream);// 创建 Excel 2003 工作簿对象  
            } else if (excel2007U.equals(fileType.trim().toLowerCase())) {  
                workbook = new XSSFWorkbook(fileStream);// 创建 Excel 2007 工作簿对象  
            } 
        }  
        return workbook;  
    }  
    
    /**
     * 找到需要插入的行数，并新建一个POI的row对象
     * @param sheet
     * @param rowIndex
     * @return
     */
 	public static Row createRow(Sheet sheet, Integer rowIndex) {
 		Row row = null;
 		if (sheet.getRow(rowIndex) != null) {
 			int lastRowNo = sheet.getLastRowNum();
 			sheet.shiftRows(rowIndex, lastRowNo, 1);
 		}
 		row = sheet.createRow(rowIndex);
 		return row;
 	}
 	/**  
 	* 判断指定的单元格是否是合并单元格  
 	* @param sheet   
 	* @param row 行下标  
 	* @param column 列下标  
 	* @return  
 	*/  
 	public static boolean isMergedRegion(Sheet sheet, int row, int column) {  
	 	int sheetMergeCount = sheet.getNumMergedRegions();  
	 	for (int i = 0; i < sheetMergeCount; i++) {  
		 	CellRangeAddress range = sheet.getMergedRegion(i);  
		 	int firstColumn = range.getFirstColumn();  
		 	int lastColumn = range.getLastColumn();  
		 	int firstRow = range.getFirstRow();  
		 	int lastRow = range.getLastRow();  
		 	if(row >= firstRow && row <= lastRow){  
		 		if(column >= firstColumn && column <= lastColumn){  
		 			return true;  
		 		}  
		 	}  
	 	}  
	 	return false;  
 	}  
 	
 	public static boolean isRowMergedRegion(Sheet sheet, XSSFCell cell) {  
	 	int sheetMergeCount = sheet.getNumMergedRegions(); 
	 	int row = cell.getRowIndex();
	 	int column = cell.getColumnIndex();
	 	for (int i = 0; i < sheetMergeCount; i++) {  
		 	CellRangeAddress range = sheet.getMergedRegion(i);  
		 	int firstColumn = range.getFirstColumn();  
		 	int lastColumn = range.getLastColumn();  
		 	int firstRow = range.getFirstRow();  
		 	int lastRow = range.getLastRow();  
		 	if(row >= firstRow && row <= lastRow){  
		 		if(column == firstColumn && column == lastColumn){  
		 			return true;  
		 		}  
		 	}  
	 	}  
	 	return false;  
 	}  
 	
 	public static boolean isColMergedRegion(Sheet sheet, XSSFCell cell) {  
	 	int sheetMergeCount = sheet.getNumMergedRegions(); 
	 	int row = cell.getRowIndex();
	 	int column = cell.getColumnIndex();
	 	for (int i = 0; i < sheetMergeCount; i++) {  
		 	CellRangeAddress range = sheet.getMergedRegion(i);  
		 	int firstColumn = range.getFirstColumn();  
		 	int lastColumn = range.getLastColumn();  
		 	int firstRow = range.getFirstRow();  
		 	int lastRow = range.getLastRow();  
		 	if(row == firstRow && row == lastRow){  
		 		if(column >= firstColumn && column <= lastColumn){  
		 			return true;  
		 		}  
		 	}  
	 	}  
	 	return false;  
 	}
 	/**  
 	* 判断指定的单元格是否是合并单元格  
 	* @param sheet   
 	* @param row 行下标  
 	* @param column 列下标  
 	* @return  
 	*/  
 	public static CellRangeAddress getMergedRegion(Sheet sheet, XSSFCell cell) {  
	 	int sheetMergeCount = sheet.getNumMergedRegions();  
	 	int row = cell.getRowIndex();
	 	int column = cell.getColumnIndex();
	 	for (int i = 0; i < sheetMergeCount; i++) {  
		 	CellRangeAddress range = sheet.getMergedRegion(i);  
		 	int firstColumn = range.getFirstColumn();  
		 	int lastColumn = range.getLastColumn();  
		 	int firstRow = range.getFirstRow();  
		 	int lastRow = range.getLastRow();  
		 	if(row >= firstRow && row <= lastRow){  
		 		if(column >= firstColumn && column <= lastColumn){  
		 			return range;
		 		}  
		 	}  
	 	}  
	 	return null;  
 	} 
 	
 	public static void removeMergedRegion(Sheet sheet, XSSFCell cell) {  
	 	int sheetMergeCount = sheet.getNumMergedRegions();  
	 	int row = cell.getRowIndex();
	 	int column = cell.getColumnIndex();
	 	for (int i = 0; i < sheetMergeCount; i++) {  
		 	CellRangeAddress range = sheet.getMergedRegion(i);  
		 	int firstColumn = range.getFirstColumn();  
		 	int lastColumn = range.getLastColumn();  
		 	int firstRow = range.getFirstRow();  
		 	int lastRow = range.getLastRow();  
		 	if(row >= firstRow && row <= lastRow){  
		 		if(column >= firstColumn && column <= lastColumn){  
		 			sheet.removeMergedRegion(i);
		 			break;
		 		}  
		 	}  
	 	}  
 	} 
 	
 	@SuppressWarnings("deprecation")
	public static String getCellValue(Cell cell){
 	    if(cell == null) return "";    
 	    if(cell.getCellType() == Cell.CELL_TYPE_STRING){    
 	        return cell.getStringCellValue();    
 	    }else if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN){    
 	        return String.valueOf(cell.getBooleanCellValue());    
 	    }else if(cell.getCellType() == Cell.CELL_TYPE_FORMULA){    
 	        return cell.getCellFormula() ;    
 	    }else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){    
 	        return String.valueOf(cell.getNumericCellValue());    
 	    }    
 	    return "";    
 	}    
 	
 	public static void makeStyle(XSSFWorkbook wb, XSSFSheet sheet, XSSFCell cell){
 		XSSFFont ztFont = wb.createFont();     
        ztFont.setItalic(true);                     // 设置字体为斜体字     
        ztFont.setColor(XSSFFont.COLOR_RED);            // 将字体设置为“红色”     
        ztFont.setFontHeightInPoints((short)22);    // 将字体大小设置为18px     
        ztFont.setFontName("华文行楷");             // 将“华文行楷”字体应用到当前单元格上     
        ztFont.setUnderline(XSSFFont.U_DOUBLE);         // 添加（Font.U_SINGLE单条下划线/Font.U_DOUBLE双条下划线）    
//      ztFont.setStrikeout(true);                  // 是否添加删除线     
        XSSFCellStyle ztStyle = (XSSFCellStyle) wb.createCellStyle();     
        ztStyle.setFont(ztFont);                    // 将字体应用到样式上面     
        cell.setCellStyle(ztStyle);               // 样式应用到该单元格上 
 	}
 	
 	/**
 	 * @param wb 
 	 * @param sheet
 	 * @param startRow 在哪行插
 	 * @param rows 插几行
 	 * */
 	public static void insertRow(XSSFWorkbook wb, XSSFSheet sheet, int startRow,int rows) {  
        startRow--;
        sheet.shiftRows(startRow+1, sheet.getLastRowNum(), rows,true,false);  
        List<int[]> mergeCellDigits = Lists.newArrayList();
        List<String> mergeCellValues = Lists.newArrayList();
        
        for (int i = 0; i < rows; i++) {  
              
              XSSFRow sourceRow = null;  
              XSSFRow targetRow = null;  
              sourceRow = sheet.getRow(startRow);  
              for (Cell c : sourceRow) {
            	XSSFCell cell = (XSSFCell) c;
            	//System.out.println(cell.getRowIndex()+1 + "," + (cell.getColumnIndex()+1) +":" +getCellValue(cell));
				if(isRowMergedRegion(sheet, cell)){
					//System.out.println("要拷贝的行中包含合并格");
					CellRangeAddress range =  getMergedRegion(sheet, cell);
					mergeCellDigits.add(new int[]{range.getFirstRow(),range.getLastRow(),range.getFirstColumn(),range.getLastColumn()});
					mergeCellValues.add(getMergedRegionValue(sheet, cell));
					removeMergedRegion(sheet, cell);
				}
			  }
              targetRow = sheet.createRow(++startRow);  
              Util.copyRow(sheet, sourceRow, targetRow);  
        }  
        
        for (int i = 0; i < mergeCellDigits.size(); i++) {
        	int[] digits = mergeCellDigits.get(i);
        	String cellValue = mergeCellValues.get(i);
        	mergeRegion(sheet, digits[0], digits[1]+rows, digits[2], digits[3]);
        	setCellValue(sheet.getRow(digits[0]).getCell(digits[2]), cellValue);
		}
        
 	}  
 	
 	public static void copyRow(XSSFWorkbook wb, XSSFSheet sheet,
			int[] sourceRows, int targetStartRow, List<RowDTO> rowList) {
 		
		int[] newSourceRows = Arrays.copyOf(sourceRows, sourceRows.length);
		//先算出被复制行的真实行号
		for (int i = 0; i < newSourceRows.length; i++) {
			for (RowDTO rowDTO : rowList) {
				if(rowDTO.getType().equals("insert")){
					if(sourceRows[i] <= rowDTO.getOrignStartRow()){
						//do nothing
					}else{
						if(rowDTO.isWriteFromStart()){
							newSourceRows[i] += rowDTO.getMyRows().size()-1 ;
						}else{
							newSourceRows[i] += rowDTO.getMyRows().size();
						}
					}
				}else if(rowDTO.getType().equals("copy")
						&& !Arrays.equals(rowDTO.getSourceRows(), newSourceRows)){
					if(sourceRows[i] <= rowDTO.getTargetStartRow()){
						//do nothing
					}else{
						newSourceRows[i] += rowDTO.getSourceRows().length;
					}
				}
			}
		}
		//开始复制
		//目标行现有内容向下移动
		int targetStartRowIndex = targetStartRow-1;
		sheet.shiftRows(targetStartRowIndex, sheet.getLastRowNum(), sourceRows.length,true,false);  
		//复制行到指定行
		for (int i = 0; i < newSourceRows.length; i++) {
			XSSFRow sourceRow = sheet.getRow(newSourceRows[i]-1);
			XSSFRow targetRow = sheet.createRow(targetStartRowIndex++);  
            Util.copyRow(sheet, sourceRow, targetRow);  
		}
	}
 	
 	public static void insertAndWriteRow(XSSFWorkbook wb, XSSFSheet sheet, int startRow, List<List<String>> myRows) {  
 		//startRow--;
 		int rows = myRows.size();
        sheet.shiftRows(startRow + 1, sheet.getLastRowNum(), rows,true,false);  
        List<int[]> mergeCellDigits = Lists.newArrayList();
        List<String> mergeCellValues = Lists.newArrayList();
        
        if(rows>0)
        for (int i = 0; i < rows; i++) {  
        	  List<String> myRow = myRows.get(i);
        	
              XSSFRow sourceRow = null;  
              XSSFRow targetRow = null;  
              sourceRow = sheet.getRow(startRow);
              for (Cell c : sourceRow) {
              	XSSFCell cell = (XSSFCell) c;
              	//System.out.println(cell.getRowIndex()+1 + "," + (cell.getColumnIndex()+1) +":" +getCellValue(cell));
  				if(isRowMergedRegion(sheet, cell)){
  					//System.out.println("要拷贝的行中包含合并格");
  					CellRangeAddress range =  getMergedRegion(sheet, cell);
  					mergeCellDigits.add(new int[]{range.getFirstRow(),range.getLastRow(),range.getFirstColumn(),range.getLastColumn()});
  					mergeCellValues.add(getMergedRegionValue(sheet, cell));
  					removeMergedRegion(sheet, cell);
  				}
  			  }
              targetRow = sheet.createRow(++startRow);  
              Util.copyRow(sheet, sourceRow, targetRow);  
              
              XSSFRow currentRow = sheet.getRow(startRow-1);
              
              
              int cellnum = 0;
              for (int j = 0; j < myRow.size(); j++) {
            	  XSSFCell cell = currentRow.getCell(cellnum);
            	  
            	  //System.out.println(cell.getRowIndex()+1 + "," + (cell.getColumnIndex()+1) +":" +getCellValue(cell));
            	  if(isColMergedRegion(sheet, cell)){
            		  //System.out.println("合并格！");
            		  CellRangeAddress range = getMergedRegion(sheet, cell);
            		  int colspan = range.getLastColumn() - range.getFirstColumn();
            		  setCellValue(cell, myRow.get(j));
            		  //System.out.println("合并格赋值"+j);
            		  cellnum += colspan + 1;
            	  }else{
            		  setCellValue(cell, myRow.get(j));
            		  //System.out.println("赋值"+j);
            		  cellnum++;
            	  }
			  }
        }
        
        for (int i = 0; i < mergeCellDigits.size(); i++) {
        	int[] digits = mergeCellDigits.get(i);
        	String cellValue = mergeCellValues.get(i);
        	mergeRegion(sheet, digits[0], digits[1]+rows, digits[2], digits[3]);
        	setCellValue(sheet.getRow(digits[0]).getCell(digits[2]), cellValue);
		}
        XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
 	}  
 	
 	
 	public static void insertAndWriteRowFromStart(XSSFWorkbook wb, XSSFSheet sheet, int startRow, List<List<String>> myRows) {  
 		//startRow--;
 		int rows = myRows.size();
 		if(rows>1){
 			sheet.shiftRows(startRow, sheet.getLastRowNum()+1, rows-1,true,false);  //因为范例行也要被修改，只用向下拉开rows-1行
 			List<int[]> mergeCellDigits = Lists.newArrayList();
 			List<String> mergeCellValues = Lists.newArrayList();
 			
 			if(rows>0)
 				for (int i = 0; i < rows; i++) {  
 					List<String> myRow = myRows.get(i);
 					
 					XSSFRow sourceRow = null;  
 					XSSFRow targetRow = null;  
 					//System.out.println("startrow:"+(startRow-1));//9/10
 					sourceRow = sheet.getRow(startRow-1);//从0开始
 					for (Cell c : sourceRow) {
 						XSSFCell cell = (XSSFCell) c;
 						//System.out.println(cell.getRowIndex()+1 + "," + (cell.getColumnIndex()+1) +":" +getCellValue(cell));
 						if(isRowMergedRegion(sheet, cell)){
 							System.out.println("要拷贝的行中包含合并格");
 							CellRangeAddress range =  getMergedRegion(sheet, cell);
 							mergeCellDigits.add(new int[]{range.getFirstRow(),range.getLastRow(),range.getFirstColumn(),range.getLastColumn()});
 							mergeCellValues.add(getMergedRegionValue(sheet, cell));
 							removeMergedRegion(sheet, cell);
 						}
 					}
 					//System.out.println("targetrow:"+startRow);
 					if(i<rows-1){
 						targetRow = sheet.createRow(startRow++);  
 						Util.copyRow(sheet, sourceRow, targetRow);  
 						for(Cell c : targetRow){ //若有公式，更新公式的脚标
 							if(StringUtils.isNotBlank(getCellFormula(c))){
 								c.setCellFormula(getFormulaAppended(c.getCellFormula()));
 							}
 						}
 						
 					}else{
 						startRow++;
 					}
 					XSSFRow currentRow = sheet.getRow(startRow-2);
 					
 					
 					int cellnum = 0;
 					for (int j = 0; j < myRow.size(); j++) {
 						XSSFCell cell = currentRow.getCell(cellnum);
 						
 						if(cell==null){
 							currentRow.createCell(cellnum).setCellValue("");
 							cell = currentRow.getCell(cellnum);
 							if(cellnum!=0){
 								cell.setCellStyle(currentRow.getCell(cellnum-1).getCellStyle());
 							}
 						}
 						//System.out.println(cell.getRowIndex()+1 + "," + (cell.getColumnIndex()+1) +":" +getCellValue(cell));
 						
 						if(isColMergedRegion(sheet, cell)){
 							//System.out.println("合并格！");
 							CellRangeAddress range = getMergedRegion(sheet, cell);
 							int colspan = range.getLastColumn() - range.getFirstColumn();
 							setCellValue(cell, myRow.get(j));
 							//System.out.println("合并格赋值"+j);
 							cellnum += colspan+1; //+ 1; 包含范例行，因此少合并一行
 						}else{
 							setCellValue(cell, myRow.get(j));
 							//System.out.println("赋值"+j);
 							cellnum++;
 						}
 					}
 				}
 			
 			for (int i = 0; i < mergeCellDigits.size(); i++) {
 				int[] digits = mergeCellDigits.get(i);
 				String cellValue = mergeCellValues.get(i);
 				mergeRegion(sheet, digits[0], digits[1]+rows-1, digits[2], digits[3]);
 				setCellValue(sheet.getRow(digits[0]).getCell(digits[2]), cellValue);
 			}
 		}else if(rows==1){
 			writeRow(wb, sheet, startRow, myRows);
 		}else{
 			//do nothing
 		}
 		XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
 	}  
 	/**   
 	* 获取合并单元格的值   
 	* @param sheet   
 	* @param row   
 	* @param column   
 	* @return   
 	*/    
 	public static String getMergedRegionValue(Sheet sheet, XSSFCell cell){    
 	    int sheetMergeCount = sheet.getNumMergedRegions();    
 	    int row = cell.getRowIndex();
	 	int column = cell.getColumnIndex();
 	    for(int i = 0 ; i < sheetMergeCount ; i++){    
 	        CellRangeAddress ca = sheet.getMergedRegion(i);    
 	        int firstColumn = ca.getFirstColumn();    
 	        int lastColumn = ca.getLastColumn();    
 	        int firstRow = ca.getFirstRow();    
 	        int lastRow = ca.getLastRow();    
 	        if(row >= firstRow && row <= lastRow){    
 	            if(column >= firstColumn && column <= lastColumn){    
 	                Row fRow = sheet.getRow(firstRow);    
 	                Cell fCell = fRow.getCell(firstColumn);    
 	                return getCellValue(fCell) ;    
 	            }    
 	        }    
 	    }    
 	        
 	    return null ;    
 	}    
 	
 	/**  
 	* 合并单元格  
 	* @param sheet   
 	* @param firstRow 开始行  
 	* @param lastRow 结束行  
 	* @param firstCol 开始列  
 	* @param lastCol 结束列  
 	*/  
 	public static void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {  
 		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));  
 	}  
 	
 	public static void writeRow(XSSFWorkbook wb, XSSFSheet sheet, int startRow, List<List<String>> myRow) {  
        int rows = myRow.size();
        //startRow--;//从1开始
        
        if(rows>0)
        for (int i = 0; i < rows; i++) {  
             XSSFRow currentRow = sheet.getRow(startRow-1 + i);  
             
             int cellnum = 0;
             for (int j = 0; j < myRow.get(i).size(); j++) {
           	  XSSFCell cell = currentRow.getCell(cellnum);
           	  if(cell==null){
					currentRow.createCell(cellnum).setCellValue("");
					cell = currentRow.getCell(cellnum);
					if(cellnum!=0){
						cell.setCellStyle(currentRow.getCell(cellnum-1).getCellStyle());
					}
				}
           	  
           	  //System.out.println(cell.getRowIndex()+1 + "," + (cell.getColumnIndex()+1) +":" +getCellValue(cell));
           	  if(isColMergedRegion(sheet, cell)){
           		  //System.out.println("合并格！");
           		  CellRangeAddress range = getMergedRegion(sheet, cell);
           		  int colspan = range.getLastColumn() - range.getFirstColumn();
           		  setCellValue(cell,myRow.get(i).get(j));
           		  //System.out.println("合并格赋值"+j);
           		  cellnum += colspan + 1;
           	  }else{
           		  setCellValue(cell,myRow.get(i).get(j));
           		  //System.out.println("赋值"+j);
           		  cellnum++;
           	  }
             }
              
        }  
        XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        
 	}
 	
 	@SuppressWarnings("deprecation")
	public static void setCellValue(XSSFCell cell, String cellValue){
 		if(cellValue!=null
 				&& !"None".equals(cellValue) && !"null".equalsIgnoreCase(cellValue)){
 			if(CommonHelper.checkNumber(cellValue)){
 				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
 				cell.setCellValue(Double.valueOf(cellValue));
 			}else if(cellValue.startsWith("#FORMULA:")){
 				cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
 				cell.setCellFormula(cellValue.replace("#FORMULA:", ""));
 			}else if(CommonHelper.Str2DateAutoFormat(cellValue)!=null){
 				cell.setCellValue(CommonHelper.Str2DateAutoFormat(cellValue));
 			}else{
 				cell.setCellType(Cell.CELL_TYPE_STRING);
 				cell.setCellValue(cellValue);
 			}
 		}
 	}

 	 /** 
     * Sheet复制 
     * @param fromSheet 
     * @param toSheet 
     * @param copyValueFlag 
     */  
    public static void copySheet(XSSFWorkbook wb,XSSFSheet fromSheet, XSSFSheet toSheet,  
            boolean copyValueFlag) {  
        //合并区域处理  
    	int pStartRow = fromSheet.getFirstRowNum();
    	int pEndRow = fromSheet.getLastRowNum();
    	int maxColumnNum = 0;  
    	
    	List<CellRangeAddress> oldRanges = new ArrayList<CellRangeAddress>();
    	for (int i = 0; i < fromSheet.getNumMergedRegions(); i++) {
    	   oldRanges.add(fromSheet.getMergedRegion(i));
    	}
    	
    	for (int i = pStartRow; i <= pEndRow; i++) {
    		//System.out.println(i);
    		XSSFRow fromRow = fromSheet.getRow(i);  
            XSSFRow toRow = toSheet.createRow(i);
            
            toRow.setHeight(fromRow.getHeight());
            
            copyRow(wb,fromSheet,toSheet,fromRow,toRow,oldRanges);
            if (fromRow.getLastCellNum() > maxColumnNum) {  
                maxColumnNum = fromRow.getLastCellNum();  
            } 
		}
    	
    	//合并一波
    	for (CellRangeAddress cellRangeAddress : oldRanges) {
    		int firstColumn = cellRangeAddress.getFirstColumn();    
  	        int lastColumn = cellRangeAddress.getLastColumn();    
  	        int firstRow = cellRangeAddress.getFirstRow();    
  	        int lastRow = cellRangeAddress.getLastRow();  
  	        CellRangeAddress newRange = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
  	        toSheet.addMergedRegion(newRange);
    	}
    	//设置列宽  
    	for (int i = 0; i < maxColumnNum; i++) { 
    		//System.out.println(fromSheet.getColumnWidth(i));
    		//System.out.println(i+":"+ fromSheet.getColumnWidthInPixels(i));
    		toSheet.setColumnWidth(i, fromSheet.getColumnWidth(i));  
        }
    	System.out.println("-----");
//    	for (int i = 0;i < maxColumnNum; i++){
//    		toSheet.setColumnWidth(i,3800);
//    	}
    	reCalculating(wb, toSheet);
    	
    } 
    
    public static void copyRow(XSSFWorkbook wb, XSSFSheet fromSheet, XSSFSheet toSheet, XSSFRow fromRow, XSSFRow toRow,  
    		List<CellRangeAddress> oldRanges){
    	Boolean copyValueFlag = true;
    	for (int i = fromRow.getFirstCellNum(); i <= fromRow.getLastCellNum(); i++) {
    		 if(i<=0){
    			 //System.out.println("空行:"+fromRow.getRowNum());
    			 continue;
    		 }
    		 XSSFCell fromCell = fromRow.getCell(i); // old cell  
             XSSFCell toCell = toRow.getCell(i); // new cell 
             if(fromCell != null){
            	 if(toCell == null){
            		 toCell = toRow.createCell(i);
//            		 if(StringUtils.isNotBlank(getCellValue(fromCell))){
//            			 setCellValue(toCell, getCellValue(fromCell));
//            		 }else{
//            			 toCell.setCellValue("1");
//            		 }
            		 CellStyle srcStyle = fromCell.getCellStyle();
            		 CellStyle newStyle = wb.createCellStyle();
            		 newStyle.cloneStyleFrom(fromCell.getCellStyle());
            		 newStyle.setFont(wb.getFontAt(srcStyle.getFontIndex()));
            	        //样式
            		 toCell.setCellStyle(newStyle);
            	        //评论
            	        if(fromCell.getCellComment() != null) {
            	        	toCell.setCellComment(fromCell.getCellComment());
            	        }
            	        // 不同数据类型处理
            	        CellType srcCellType = fromCell.getCellTypeEnum();
            	        toCell.setCellType(srcCellType);
            	        if(copyValueFlag) {
            	            if(srcCellType == CellType.NUMERIC) {
            	                if(DateUtil.isCellDateFormatted(fromCell)) {
            	                	toCell.setCellValue(fromCell.getDateCellValue());
            	                } else {
            	                	toCell.setCellValue(fromCell.getNumericCellValue());
            	                }
            	            } else if(srcCellType == CellType.STRING) {
            	            	toCell.setCellValue(fromCell.getRichStringCellValue());
            	            } else if(srcCellType == CellType.BLANK) {

            	            } else if(srcCellType == CellType.BOOLEAN) {
            	            	toCell.setCellValue(fromCell.getBooleanCellValue());
            	            } else if(srcCellType == CellType.ERROR) {
            	            	toCell.setCellErrorValue(fromCell.getErrorCellValue());
            	            } else if(srcCellType == CellType.FORMULA) {
            	            	toCell.setCellFormula(fromCell.getCellFormula());
            	            } else {
            	            }
            	        }
            	 }
             }
		}
    }
    
    public static void copyCell(Workbook wb, Cell srcCell, Cell distCell, boolean copyValueFlag) {
        CellStyle newStyle = wb.createCellStyle();
        CellStyle srcStyle = srcCell.getCellStyle();

        newStyle.cloneStyleFrom(srcStyle);
        newStyle.setFont(wb.getFontAt(srcStyle.getFontIndex()));
        //样式
        distCell.setCellStyle(newStyle);
        //评论
        if(srcCell.getCellComment() != null) {
            distCell.setCellComment(srcCell.getCellComment());
        }
        // 不同数据类型处理
        CellType srcCellType = srcCell.getCellTypeEnum();
        distCell.setCellType(srcCellType);
        if(copyValueFlag) {
            if(srcCellType == CellType.NUMERIC) {
                if(DateUtil.isCellDateFormatted(srcCell)) {
                    distCell.setCellValue(srcCell.getDateCellValue());
                } else {
                    distCell.setCellValue(srcCell.getNumericCellValue());
                }
            } else if(srcCellType == CellType.STRING) {
                distCell.setCellValue(srcCell.getRichStringCellValue());
            } else if(srcCellType == CellType.BLANK) {

            } else if(srcCellType == CellType.BOOLEAN) {
                distCell.setCellValue(srcCell.getBooleanCellValue());
            } else if(srcCellType == CellType.ERROR) {
                distCell.setCellErrorValue(srcCell.getErrorCellValue());
            } else if(srcCellType == CellType.FORMULA) {
            	String srcFormula = srcCell.getCellFormula();
                distCell.setCellFormula(srcCell.getCellFormula());
            } else {
            }
        }
        
    }
    
   public static void reCalculating(XSSFWorkbook wb, XSSFSheet sheet){
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        for(Row r : sheet) {
           for(Cell c : r) {
                if(c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    evaluator.evaluateFormulaCell(c);
                }
           }
        }
    }
   
   /*****sheet导出HTML****/
   public String readExcelToHtml(XSSFWorkbook wb, XSSFSheet sheet, boolean isWithStyle){
	   StringBuffer sb = new StringBuffer();  
	   return sb.toString();
   }
   
   public static String getCellFormula(Cell cell){
	   if(CellType.FORMULA == cell.getCellTypeEnum()){
		   return cell.getCellFormula();
	   }else{
		   return null;
	   } 
   }
   
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
	
	
	public static String getFormulaAppended(String input){
		// 创建一个正则表达式模式，用以匹配一个单词（\b=单词边界）
		  String patt = "\\b[A-Z]{1}\\d+\\b";
		  Pattern r = Pattern.compile(patt);
		  Matcher m = r.matcher(input);
		  StringBuffer sb = new StringBuffer();
		  int count = 0;
		  int start = 0;
		  while(m.find()){
			  String word = m.group();
			  System.out.println(word);
			  if(count==0){
				  sb.append(input.substring(start, m.start()));
				  sb.append(getCharEn(word));
				  sb.append(getCharNum(word) + 1);
				  start = m.end();
			  }else{
				  sb.append(input.substring(start, m.start()));
				  sb.append(getCharEn(word));
				  sb.append(getCharNum(word) + 1);
				  start = m.end();
			  }
			  count++;
		  }
		  sb.append(input.substring(start));
		  return sb.toString();
	}
}
