package com.kev.team.centermall.order.common.excel.entity;
import java.util.Map;
/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class RuturnPropertyParam {
	/***
	 * property name in java bean
	 */
	private String name = null;
	
	/***
	 * sequeence in excel
	 */
	private String column = null;
	/***
	 * excel title，not data
	 */
	private String excelTitleName = null;
	/***
	 * data type
	 */
	private String dataType = null;
	private String maxLength = null;
//	private String fixity = null;
//	private String codeTableName = null;
	/***
	 * date format ,if property type is "Date"
	 */
	private String dateFormat=null;
	/***
	 * default value is java bean
	 */
	private String defaultValue = null;
	private boolean isConvertable=false;
	Map<String,String> convertMap=null;
	public String getColumn() {
		return column;
	}
	public void setColumn(String column) {
		this.column = column;
	}
	public String getDataType() {
		return dataType;
	}
	public void setDataType(String dataType) {
		this.dataType = dataType;
	}
	public String getDefaultValue() {
		return defaultValue;
	}
	public void setDefaultValue(String defaultValue) {
		this.defaultValue = defaultValue;
	}
	public String getExcelTitleName() {
		return excelTitleName;
	}
	public void setExcelTitleName(String excelTitleName) {
		this.excelTitleName = excelTitleName;
	}
	public String getMaxLength() {
		return maxLength;
	}
	public void setMaxLength(String maxLength) {
		this.maxLength = maxLength;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public boolean isConvertable() {
		return isConvertable;
	}
	public void setConvertable(boolean isConvertable) {
		this.isConvertable = isConvertable;
	}
	public Map<String, String> getConvertMap() {
		return convertMap;
	}
	public void setConvertMap(Map<String, String> convertMap) {
		this.convertMap = convertMap;
	}
	public String getDateFormat() {
		return dateFormat;
	}
	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}
	
	
	
}
