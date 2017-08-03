package com.kev.team.centermall.order.common.excel.config;

import java.util.HashMap;
import java.util.Map;

/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class RuturnConfig {
	private String className = null;

	private Map propertyMap = new HashMap();
	
	public String getClassName() {
		return className;
	}

	public void setClassName(String className) {
		this.className = className;
	}

	public Map getPropertyMap() {
		return propertyMap;
	}

	public void setPropertyMap(Map propertyMap) {
		this.propertyMap = propertyMap;
	}
	
	
}
