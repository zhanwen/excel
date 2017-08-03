package com.kev.team.centermall.order.common.excel.file;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.springframework.beans.BeanWrapper;
import org.springframework.beans.BeanWrapperImpl;

import com.kev.team.centermall.order.common.excel.config.RuturnConfig;
import com.kev.team.centermall.order.common.excel.entity.RuturnPropertyParam;
import com.string.widget.util.ValueWidget;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class ExcelToModelImpl {
	private InputStream inputStream = null;

	private File excelFile = null;
	
	private RuturnConfig excelConfig = null;

	public ExcelToModelImpl(InputStream inputStream, RuturnConfig excelConfig) {
		this.excelConfig = excelConfig;
		this.inputStream = inputStream;
	}
	
	public ExcelToModelImpl(File excelFile, RuturnConfig excelConfig) {
		this.excelConfig = excelConfig;
		this.excelFile = excelFile;
	}

	public List getModelList() throws ParseException {
		List modelList = new ArrayList();
		try {
			Workbook book;

			book = Workbook.getWorkbook(this.inputStream);
			Sheet sheet = book.getSheet(0);

			// editBook = Workbook.createWorkbook(file, book);

			// System.out.println("rows = " + sheet.getRows() + "  cols ="+
			// sheet.getColumns());
			for (int i = 1; i < sheet.getRows(); i++) {
				Object obj = this.getModelInstance(excelConfig.getClassName());// new
																			// a
																			// object
				BeanWrapper bw = new BeanWrapperImpl(obj);
				// ��excelÿһ�е�ֵ����ȡֵ.
				for (int j = 0; j < sheet.getColumns(); j++) {

					// System.out.println("i = " + i + " j =" + j);
					// Excel title name
					String excelTitleName = sheet.getCell(j, 0).getContents()
							.trim();
					// ȡ��Excelֵ
					String value = sheet.getCell(j, i).getContents().trim();
					//
					RuturnPropertyParam propertyBean = (RuturnPropertyParam) excelConfig
							.getPropertyMap().get(excelTitleName);

					// System.out.println("i = " + i + "  j =" + j + " value ="+
					// value + " title = " + excelTitleName);

					if (propertyBean != null) {// is null when is in excel ,but
												// not in bean
//						System.out.println("propertyName = "
//								+ propertyBean.getName());
						//should be in front of convert map
						if (value == null || value.length() < 1) {
							value = propertyBean.getDefaultValue();
						}
						
						if (propertyBean.isConvertable()) {
							value = propertyBean.getConvertMap().get(value);
						}
						
						String dateType = propertyBean.getDataType();
						Date date2 = null;
						//convert string to Date
						if (dateType.equalsIgnoreCase("Date")) {
							String dateFormat = propertyBean.getDateFormat();
							if (ValueWidget.isHasValue(dateFormat)) {
								SimpleDateFormat sdf = new SimpleDateFormat(
										dateFormat);
								date2 = sdf.parse(value);
							}
						}
						if (date2 == null) {
							bw.setPropertyValue(propertyBean.getName(), value);
						} else {//property type is Date
							bw.setPropertyValue(propertyBean.getName(), date2);
						}

					}

				}
				modelList.add(obj);
			}
			book.close();

		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return modelList;
	}
	
	public List getModelListByFile() throws ParseException {
		List modelList = new ArrayList();
		try {
			Workbook book;

			book = Workbook.getWorkbook(this.excelFile);
			Sheet sheet = book.getSheet(0);

			// editBook = Workbook.createWorkbook(file, book);

			// System.out.println("rows = " + sheet.getRows() + "  cols ="+
			// sheet.getColumns());
			for (int i = 1; i < sheet.getRows(); i++) {
				Object obj = this.getModelInstance(excelConfig.getClassName());// new
																			// a
																			// object
				BeanWrapper bw = new BeanWrapperImpl(obj);
				// ��excelÿһ�е�ֵ����ȡֵ.
				for (int j = 0; j < sheet.getColumns(); j++) {

					// System.out.println("i = " + i + " j =" + j);
					// Excel title name
					String excelTitleName = sheet.getCell(j, 0).getContents()
							.trim();
					// ȡ��Excelֵ
					String value = sheet.getCell(j, i).getContents().trim();
					//
					RuturnPropertyParam propertyBean = (RuturnPropertyParam) excelConfig
							.getPropertyMap().get(excelTitleName);

					// System.out.println("i = " + i + "  j =" + j + " value ="+
					// value + " title = " + excelTitleName);

					if (propertyBean != null) {// is null when is in excel ,but
												// not in bean
//						System.out.println("propertyName = "
//								+ propertyBean.getName());
						//should be in front of convert map
						if (value == null || value.length() < 1) {
							value = propertyBean.getDefaultValue();
						}
						
						if (propertyBean.isConvertable()) {
							value = propertyBean.getConvertMap().get(value);
						}
						
						String dateType = propertyBean.getDataType();
						Date date2 = null;
						//convert string to Date
						if (dateType.equalsIgnoreCase("Date")) {
							String dateFormat = propertyBean.getDateFormat();
							if (ValueWidget.isHasValue(dateFormat)) {
								SimpleDateFormat sdf = new SimpleDateFormat(
										dateFormat);
								date2 = sdf.parse(value);
							}
						}
						if (date2 == null) {
							bw.setPropertyValue(propertyBean.getName(), value);
						} else {//property type is Date
							bw.setPropertyValue(propertyBean.getName(), date2);
						}

					}

				}
				modelList.add(obj);
			}
			book.close();

		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return modelList;
	}

	/***
	 * Instantiate object
	 * 
	 * @param className
	 * @return
	 */
	private Object getModelInstance(String className) {
		Object obj = null;
		try {
			obj = Class.forName(className).newInstance();
//			System.out.println("init Class = " + className);
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		}
		return obj;
	}
}
