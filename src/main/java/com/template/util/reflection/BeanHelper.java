package com.template.util.reflection;


import java.beans.BeanInfo;
import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;


public class BeanHelper {

	private static SimpleDateFormat longStrTime=new SimpleDateFormat("yyyy-MM-dd HH:mm:SS");


	/**
	 * source的非空同名属性拷贝到target中去 (从左拷到右)
	 * @param source
	 * @param target
	 */
	public static void copyPropertiesIgnoreNull(Object source, Object target) {
//		// 获取一个Bean的非空属性
//		final BeanWrapper src = new BeanWrapperImpl(source);
//		PropertyDescriptor[] pds = src.getPropertyDescriptors();
//		Set<String> emptyNames = new HashSet<String>();
//		for (PropertyDescriptor pd : pds) {
//			Object srcValue = src.getPropertyValue(pd.getName());
//			if (srcValue == null)
//				emptyNames.add(pd.getName());
//		}
//		String[] result = new String[emptyNames.size()];
//		String[] nullPropertyNames = emptyNames.toArray(result);
//		BeanUtils.copyProperties(source, target, nullPropertyNames);
	}



	/**
	 * map对象转object
	 * @param source
	 * @param target
	 */
	public static void convertMapToObject(Map<String, Object> source, Object target) {
//		ConvertUtils.register(new Converter() {
//			public Object convert(Class type, Object value) {
//				if (value instanceof Date) {
//					return value;
//				}
//				SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:SS");
//				try {
//					return simpleDateFormat.parse(value.toString());
//				} catch (ParseException e) {
//					e.printStackTrace();
//				}
//				return null;
//			}
//		}, Date.class);
//		try {
//			org.apache.commons.beanutils.BeanUtils.populate(target, source);
//		} catch (IllegalAccessException | InvocationTargetException e) {
//			e.printStackTrace();
//		}
	}



	/**
	 * 对象转map
	 * @param obj
	 * @return map
	 * @throws InvocationTargetException
	 * @throws IllegalAccessException
	 * @throws IntrospectionException
	 */
	public static Map<String, Object> convertBeanToMap(Object obj) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
		if(obj == null){
			return null;
		}
		Map<String, Object> map = new HashMap<String, Object>();
		BeanInfo beanInfo = Introspector.getBeanInfo(obj.getClass());
		PropertyDescriptor[] propertyDescriptors = beanInfo.getPropertyDescriptors();
		for (PropertyDescriptor property : propertyDescriptors) {
			String key = property.getName();
			// 过滤class属性
			if (!key.equals("class")) {
				// 得到property对应的getter方法
				Method getter = property.getReadMethod();
				Object value = getter.invoke(obj);
				map.put(key, value);
			}
		}
		return map;
	}





	//对象转字符串
	public static String convertObjToString(Object object){
		if (object!=null){
			if (object instanceof Date){
				return longStrTime.format(object);
			}
			return object.toString();
		}
		return null;
	}


	//map数据转

	public static void main(String[] args) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
//		Map<String, Object> map = convertBeanToMap(new Template());
//		System.out.println(map);
	}



}
