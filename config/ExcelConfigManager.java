package com.kev.team.centermall.order.common.excel.config;

import org.dom4j.Element;

/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public interface ExcelConfigManager {
	public Element getModelElement(String modelName);
	public RuturnConfig getModel(String modelName);
}
