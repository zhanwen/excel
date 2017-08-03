package com.kev.team.centermall.order.common.excel.config;
/**
 * 
 * @author Hanwen
 * @date 2017年7月21日
 * @version 1.0.0
 */
public class ConfigConstant {
	
	/*
	 * model 的查找编号  id
	 */
	public static final String MODEL_ID = "id";
	
	/*
	 * 从Excel 到 Model 转换的目标 javabean class
	 */
	public static final String MODEL_CLASS="class";
	
	/*
	 * excel ���ж�Ӧ��javabean�е� javabean�������
	 */
	public static final String PROPERTY_NAME = "name";
	
	/*
	 * column start from one
	 */
	public static final String PROPERTY_CLOUMN="column";
	
	/*
	 * excel ���ж�Ӧ��javabean�е� excel�б������  (ȡֵ��ʱ��,������ javabean���Ժ�Excel�еı����Ӧȡֵ)
	 */
	public static final String PROPERTY_EXCEL_TITLE_NAME = "excelTitleName";
	
	/*
	 * excel�е������Ҫ��ת�����������
	 */
	public static final String PROPERTY_DATA_TYPE = "dataType";
	
	/*
	 * excel�е������󳤶�
	 */
	public static final String PROPERTY_MAX_LENGTH="maxLength";
	/***
	 *  whether can be convert
	 */
	public static final String PROPERTY_ISCONVERTABLE="isConvertable";
	/*
	 * ��excel��û�������,��Ҫϵͳ��javabean�����е�ĳһ��ֵ����һ����̬�����ֵ(ָ����JavaBean���������ֵ������ͳһ����Ĺ̶�ֵ).����Щ����,��������Ĭ��ֵ
	 * 
	 * ���ֵ����Ϊ fixity = "yes" ,Ĭ��Ϊ no.
	 */
//	public static final String PROPERTY_FIXITY="fixity";
	
	/*
	 * �ڵ���excelʱ,��Щֵ�ڴ���ϵͳʱ,ʹ�õ�ֵ��Ҫת��
	 * ��:һ���������޵������б�����{����(Y),����(C),����(D)}
	 * ϵͳ�д����ֻ���� Y,C,D֮���ֵ,����Excel��ֵȴ������,����,���������ֵ,��Ҫת��.
	 * ת����ʽ:ȡ��Excel�еľ���ֵ,���ϱ�������ǰ�,�Ӵ����Map��ȡֵ.�� ȡֵ����  C = Map.get("bgqx����"); 
	 * 
	 */
//	public static final String PROPERTY_CODE_TABLE_NAME="codeTableName";
	/*
	 * ���ֵΪ��,���õ�Ĭ��ֵ
	 */
	public static final String PROPERTY_DEFAULT="default";
	public static final String PROPERTY_FORMAT="format";
}
