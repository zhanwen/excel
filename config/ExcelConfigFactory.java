package com.kev.team.centermall.order.common.excel.config;

import com.kev.team.centermall.order.common.excel.config.impl.ExcelConfigManagerImpl;
/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class ExcelConfigFactory {
	private static ExcelConfigManager instance = new ExcelConfigManagerImpl();
	public static ExcelConfigManager createExcelConfigManger(){
		return instance;
	}
}
