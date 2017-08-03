package com.kev.team.centermall.order.common.excel.file;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.BeanWrapper;
import org.springframework.beans.BeanWrapperImpl;

import com.common.util.SystemUtil;
import com.kev.team.centermall.order.common.excel.config.ExcelConfigFactory;
import com.kev.team.centermall.order.common.excel.config.ExcelConfigManager;
import com.kev.team.centermall.order.common.excel.config.RuturnConfig;
import com.kev.team.centermall.order.common.excel.entity.RuturnPropertyParam;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class ModelToExcelUtil {
	
	public static void model2Excel(HttpServletResponse response, String modelName, List models, String fileName)
			throws IOException, RowsExceededException, WriteException {
		
		//read "ImportExcelToModel.xml" file
		ExcelConfigManager configManager = ExcelConfigFactory
				.createExcelConfigManger();
		RuturnConfig returnConfig = configManager.getModel(modelName);
		response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xls");
		WritableWorkbook wb = Workbook.createWorkbook(response.getOutputStream());
		WritableSheet wsheet = wb.createSheet("first page", 0);

		WritableFont font1 = new WritableFont(WritableFont.ARIAL, 10,
				WritableFont.BOLD);//Set the table header in bold
		WritableCellFormat format1 = new WritableCellFormat(font1);
		Map propertyMap = returnConfig.getPropertyMap();
		int columns_size = propertyMap.size();
		
		//Setting  excel table header sequence according column in xml;
		//column start from one
		for (Iterator it = propertyMap.keySet().iterator(); it.hasNext();) {
			String key = (String) it.next();
			RuturnPropertyParam modelProperty = (RuturnPropertyParam) propertyMap
					.get(key);
			int column = Integer.parseInt(modelProperty.getColumn());//sequence  in excel file
			wsheet.addCell(new Label(column - 1, 0, key, format1));
		}

		//write real data to excel file from beans list
		for (int i = 0; i < models.size(); i++) {
			Object obj = models.get(i);
			BeanWrapper bw = new BeanWrapperImpl(obj);//use spring tools(java reflection)
			for (int k = 0; k < columns_size; k++) {
				String excelTitleName = wsheet.getCell(k, 0).getContents()
						.trim();//title name in excel
				RuturnPropertyParam modelProperty = (RuturnPropertyParam) propertyMap
						.get(excelTitleName);
				String beanPropertyName = modelProperty.getName();//property name in java object
				Object propertyValue =  bw
						.getPropertyValue(beanPropertyName);//property value in java object
				
				String dataType=modelProperty.getDataType();
				if(dataType.equalsIgnoreCase("Date")){//convert date to string
					SimpleDateFormat sdf = new SimpleDateFormat(modelProperty.getDateFormat());
					if(propertyValue == null) {
						propertyValue = modelProperty.getDefaultValue();
					}else {
						propertyValue= sdf.format((Date)propertyValue);
					}
				}
				if(propertyValue == null) {
					propertyValue = modelProperty.getDefaultValue();
				}
				if(modelProperty.isConvertable()){//whether is convertable ,see xml file
					Map<String,String> convertMap=modelProperty.getConvertMap();
					propertyValue=(String) SystemUtil.reverseMap(convertMap).get(propertyValue);
				}
				
				int setColumn = Integer.parseInt(modelProperty.getColumn()) - 1;
				
				switch (setColumn) {
					case 0:  wsheet.setColumnView(setColumn, 23);break;
					case 1:  wsheet.setColumnView(setColumn, 20);break;
					case 2:  wsheet.setColumnView(setColumn, 20);break;
					case 3:  wsheet.setColumnView(setColumn, 20);break;
					case 4:  wsheet.setColumnView(setColumn, 23);break;
					case 5:  wsheet.setColumnView(setColumn, 10);break;
					case 6:  wsheet.setColumnView(setColumn, 70);break;
					case 7:  wsheet.setColumnView(setColumn, 10);break;
					case 8:  wsheet.setColumnView(setColumn, 10);break;
					case 9:  wsheet.setColumnView(setColumn, 10);break;
					case 10: wsheet.setColumnView(setColumn, 15);break;
					case 11: wsheet.setColumnView(setColumn, 15);break;
					case 12: wsheet.setColumnView(setColumn, 70);break;
					case 13: wsheet.setColumnView(setColumn, 20);break;
					case 14: wsheet.setColumnView(setColumn, 20);break;
				}
				
				wsheet.addCell(new Label(Integer.parseInt(modelProperty
						.getColumn()) - 1, i + 1, String.valueOf(propertyValue)));
			}
		}
		wb.write();//save excel
		wb.close();//close excel file
	}
	
	/***
	 * export to excel file from java benas
	 * convert java beans to excel file
	 * 
	 * @param excelFile
	 * @param modelName
	 * @param models
	 * @return
	 * @throws IOException
	 * @throws RowsExceededException
	 * @throws WriteException
	 */
	public static File model2Excel(File excelFile, String modelName, List models)
			throws IOException, RowsExceededException, WriteException {
		if (excelFile.exists()) {//delete excel file if this file has exist
			excelFile.delete();
		}
		if (!excelFile.exists()) {
			excelFile.createNewFile();//create a new excel file
		}
		//read "ImportExcelToModel.xml" file
		ExcelConfigManager configManager = ExcelConfigFactory
				.createExcelConfigManger();
		RuturnConfig returnConfig = configManager.getModel(modelName);
		WritableWorkbook wb = Workbook.createWorkbook(excelFile);
		WritableSheet wsheet = wb.createSheet("first page", 0);

		WritableFont font1 = new WritableFont(WritableFont.ARIAL, 10,
				WritableFont.BOLD);//Set the table header in bold
		WritableCellFormat format1 = new WritableCellFormat(font1);
		Map propertyMap = returnConfig.getPropertyMap();
		int columns_size = propertyMap.size();
		
		//Setting  excel table header sequence according column in xml;
		//column start from one
		for (Iterator it = propertyMap.keySet().iterator(); it.hasNext();) {
			String key = (String) it.next();
			RuturnPropertyParam modelProperty = (RuturnPropertyParam) propertyMap
					.get(key);
			int column = Integer.parseInt(modelProperty.getColumn());//sequence  in excel file
			wsheet.addCell(new Label(column - 1, 0, key, format1));
		}

		//write real data to excel file from beans list
		for (int i = 0; i < models.size(); i++) {
			Object obj = models.get(i);
			BeanWrapper bw = new BeanWrapperImpl(obj);//use spring tools(java reflection)
			for (int k = 0; k < columns_size; k++) {
				String excelTitleName = wsheet.getCell(k, 0).getContents()
						.trim();//title name in excel
				RuturnPropertyParam modelProperty = (RuturnPropertyParam) propertyMap
						.get(excelTitleName);
				String beanPropertyName = modelProperty.getName();//property name in java object
				Object propertyValue =  bw
						.getPropertyValue(beanPropertyName);//property value in java object
				
				String dataType=modelProperty.getDataType();
				if(dataType.equalsIgnoreCase("Date")){//convert date to string
					SimpleDateFormat sdf = new SimpleDateFormat(modelProperty.getDateFormat());
					propertyValue= sdf.format((Date)propertyValue);
				}
				if(modelProperty.isConvertable()){//whether is convertable ,see xml file
					Map<String,String> convertMap=modelProperty.getConvertMap();
					propertyValue=(String) SystemUtil.reverseMap(convertMap).get(propertyValue);
				}
				
				wsheet.addCell(new Label(Integer.parseInt(modelProperty
						.getColumn()) - 1, i + 1, (String)propertyValue));
			}
		}
		wb.write();//save excel
		wb.close();//close excel file
		return excelFile;
	}

	/***
	 * 
	 * @param excelFileStr   :path of excel file;type:String
	 * @param modelName
	 * @param models
	 * @return
	 * @throws IOException
	 * @throws RowsExceededException
	 * @throws WriteException
	 */
	public static File model2Excel(String excelFileStr, String modelName, List models)
			throws IOException, RowsExceededException, WriteException {
		File excelFile=new File(excelFileStr);
		return model2Excel(excelFile, modelName, models);
	}
}
