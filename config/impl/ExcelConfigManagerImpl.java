package com.kev.team.centermall.order.common.excel.config.impl;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import com.kev.team.centermall.order.common.excel.config.ConfigConstant;
import com.kev.team.centermall.order.common.excel.config.ExcelConfigManager;
import com.kev.team.centermall.order.common.excel.config.RuturnConfig;
import com.kev.team.centermall.order.common.excel.entity.RuturnPropertyParam;
import com.kev.team.centermall.order.common.excel.util.ExcelHWUtil;
import com.string.widget.util.ValueWidget;

/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class ExcelConfigManagerImpl implements ExcelConfigManager {

	private String configName = "ExcelModeMappingl.xml";
	private SAXReader saxReader;
	private Document doc;
	private Element root;

	/****
	 * read ImportExcelToModel.xml
	 */
	public ExcelConfigManagerImpl() {
		String str = this.getClass().getResource("/").getPath();
		// this.getClass().getResourceAsStream("/");

		// InputStream in = ClassLoader.getSystemResourceAsStream(configName);

		InputStream in = Thread.currentThread().getContextClassLoader()
				.getResourceAsStream(configName);

		saxReader = new SAXReader();

		try {
			// doc = saxReader.read(new File(str + configName));
			doc = saxReader.read(in);
		} catch (DocumentException e) {
			e.printStackTrace();
		}
		root = doc.getRootElement();
	}

	/***
	 * get node through class name
	 */
	public Element getModelElement(String modelName) {
		List list = root.elements();
		Element model = null;
		Element returnModel = null;
		for (Iterator it = list.iterator(); it.hasNext();) {
			model = (Element) it.next();
			if (model.attributeValue("id").equals(modelName)) {
				returnModel = model;
				break;
			}
		}
		return returnModel;
	}

	public RuturnConfig getModel(String modelName) {
		Element model = this.getModelElement(modelName);
		RuturnConfig result = new RuturnConfig();
		if (model != null) {
			result.setClassName(model
					.attributeValue(ConfigConstant.MODEL_CLASS));
			result.setPropertyMap(this.getPropertyMap(model));
		}
		return result;
	}

	/****
	 * Analytic/parse properties of the class
	 * 
	 * @param model
	 * @return
	 */
	private Map<String, RuturnPropertyParam> getPropertyMap(Element model) {
		Map<String, RuturnPropertyParam> propertyMap = new HashMap<String, RuturnPropertyParam>();
		List list = model.elements();// An attribute of a class
		Element property = null;
		for (Iterator it = list.iterator(); it.hasNext();) {
			property = (Element) it.next();
			RuturnPropertyParam modelProperty = new RuturnPropertyParam();
			modelProperty.setName(property
					.attributeValue(ConfigConstant.PROPERTY_NAME));//property name in java bean
			
			modelProperty.setColumn(property
					.attributeValue(ConfigConstant.PROPERTY_CLOUMN));//sequeence in excel
			modelProperty.setExcelTitleName(property
					.attributeValue(ConfigConstant.PROPERTY_EXCEL_TITLE_NAME));//table title in excel
			
			modelProperty.setDataType(property
					.attributeValue(ConfigConstant.PROPERTY_DATA_TYPE));//data type:[String,Date]
			
			modelProperty.setMaxLength(property
					.attributeValue(ConfigConstant.PROPERTY_MAX_LENGTH));
			modelProperty.setColumn(property.attributeValue(ConfigConstant.PROPERTY_CLOUMN));
			
			//if data type is "Date"
			modelProperty.setDateFormat(property.attributeValue(ConfigConstant.PROPERTY_FORMAT));
			String isConvertableStr = property
					.attributeValue(ConfigConstant.PROPERTY_ISCONVERTABLE);
			boolean isConvertable = Boolean.parseBoolean(isConvertableStr);
			modelProperty.setConvertable(isConvertable);
			if (isConvertable) {
				List map_list = property.elements();
				Element property_tmp = null;
				for (Iterator it2 = map_list.iterator(); it2.hasNext();) {
					property_tmp = (Element) it2.next();
					if (property_tmp != null) {
						if (property_tmp.getName().equals("map")) {
							Map<String, String> map_tmp = new HashMap<String, String>();
							List entities = property_tmp.elements();
							for (int i = 0; i < entities.size(); i++) {
								Element entity = (Element) entities.get(i);
								String key2 = entity.attributeValue(ExcelHWUtil.MAP_KEY);
								String value2 = entity.attributeValue(ExcelHWUtil.MAP_VALUE);
								if (ValueWidget.isHasValue(key2)) {
									map_tmp.put(key2, value2);
								}
							}
							modelProperty.setConvertMap(map_tmp);

						}
					}
				}
			}

			// modelProperty.setFixity(property.attributeValue(ConfigConstant.PROPERTY_FIXITY));
			// modelProperty.setCodeTableName(property.attributeValue(ConfigConstant.PROPERTY_CODE_TABLE_NAME));
			modelProperty.setDefaultValue(property
					.attributeValue(ConfigConstant.PROPERTY_DEFAULT));//default value


			propertyMap.put(modelProperty.getExcelTitleName(), modelProperty);

		}
		return propertyMap;
	}

	public static void main(String[] args) {
		ExcelConfigManagerImpl test = new ExcelConfigManagerImpl();
		test.getModel("deptModel");
		// System.out.println(model.attributeValue("class"));

	}
}
