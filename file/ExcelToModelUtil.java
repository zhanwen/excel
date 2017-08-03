package com.kev.team.centermall.order.common.excel.file;
import java.io.File;
import java.io.InputStream;
import java.text.ParseException;
import java.util.List;

import com.kev.team.centermall.order.common.excel.config.ExcelConfigFactory;
import com.kev.team.centermall.order.common.excel.config.ExcelConfigManager;
/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class ExcelToModelUtil {
	
	public static List getModelList(InputStream inputStream,String modelName) throws ParseException{
		ExcelConfigManager configManager = ExcelConfigFactory.createExcelConfigManger();
		ExcelToModelImpl etm = null;
		try {
			etm = new ExcelToModelImpl(inputStream,configManager.getModel(modelName));
		} catch (Exception e) {
			e.printStackTrace();
		}
		List modelList = etm.getModelList();
		return modelList;
	}
	
	/***
	 * 
	 * @param excelFile
	 * @param modelName   : id in excel config file
	 * @return
	 * @throws ParseException 
	 */
	public static List getModelList(File excelFile,String modelName) throws ParseException{
		if(!excelFile.exists()){
			System.out.println( ExcelToModelUtil.class.getSimpleName()+" [getModelList:]"+excelFile.getAbsolutePath()+" does not exist.");
			return null;
		}
		ExcelConfigManager configManager = ExcelConfigFactory.createExcelConfigManger();
		ExcelToModelImpl etm =  new ExcelToModelImpl(excelFile,configManager.getModel(modelName));
		List modelList = etm.getModelList();
		
		return modelList;
	}
	public static List getModelList(String excelFileStr,String modelName) throws ParseException{
		File excelFile=new File(excelFileStr);
		return getModelList(excelFile, modelName);
	}
}
