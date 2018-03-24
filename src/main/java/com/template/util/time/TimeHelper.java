package com.template.util.time;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class TimeHelper {
    // 获取当前日期所在周的上下界  
    public final static int CURRENT_WEEK = 0;  
    // 获取当前日期上一周的上下界  
    public final static int LAST_WEEK = 1;  
    // 获取当前日期所在月的上下界  
    public final static int CURRENT_MONTH = 2;  
    // 获取当前日期上一月的上下界  
    public final static int LAST_MONTH = 3;  
    // 获取当前日期所在月的上下界  
    public final static int CURRENT_YEAR = 4;  
    // 获取当前日期上一月的上下界  
    public final static int LAST_YEAR = 5;  
	
	
	
	/**
	 * 将时间搓转成时间
	 * 
	 * @param l
	 * @return
	 * @param id
	 *            String 唯一标识
	 * @param personid
	 *            String 用户唯一标识
	 * @return BaseReturn 基本返回对象
	 * @变更记录 2017年3月2日 上午9:23:12 陈树宏创建
	 */
	public static Date longSetDate(Long l) {
		Date date = new Date(l);
		return date;
	}

	/**
	 * 计算目标时间与当前时间距离（毫秒）
	 * 
	 * @param targetHour
	 * @return
	 * @param id
	 *            String 唯一标识
	 * @param personid
	 *            String 用户唯一标识
	 * @return BaseReturn 基本返回对象
	 * @变更记录 2017年3月9日 下午6:56:05 陈仁浩创建
	 */
	public static long getTimeDelay(int targetHour, int targetMinute, int targetSecond) {
		long delay = 0;
		Calendar targetTime = Calendar.getInstance();
		targetTime.set(Calendar.HOUR_OF_DAY, targetHour);
		targetTime.set(Calendar.MINUTE, targetMinute);

		Calendar rightNow = Calendar.getInstance();
		int currentHour = rightNow.get(Calendar.HOUR_OF_DAY);
		int currentMinute = rightNow.get(Calendar.MINUTE);
		int currentSecond = rightNow.get(Calendar.SECOND);

		// 判断当前时间在目标时间之前还是之后

		if (targetTime.compareTo(rightNow) == 0) {
			delay = 0;
		} else if (targetTime.compareTo(rightNow) > 0) {
			delay = targetTime.getTimeInMillis() - rightNow.getTimeInMillis();
		} else {
			// 明天再启动
			delay = targetTime.getTimeInMillis() + 24 * 60 * 60 * 1000 - rightNow.getTimeInMillis();
		}
		return delay;
	}

	/**
	 * 传入某日的任意时间，返回该天的开始和结束时间 用于数据库查询 查询某天的数据
	 * 
	 * @param date
	 * @return
	 * @param id
	 *            String 唯一标识
	 * @param personid
	 *            String 用户唯一标识
	 * @return BaseReturn 基本返回对象
	 * @变更记录 2017年3月10日 下午7:11:14 陈仁浩创建
	 */
	public static Date[] getDayBeginAndEnd(Date date) {
		Date dayFrom = new Date(date.getTime());
		Date dayTo = new Date(date.getTime());
		dayFrom.setHours(0);
		dayFrom.setMinutes(0);
		dayFrom.setSeconds(0);
		dayTo.setHours(23);
		dayTo.setMinutes(59);
		dayTo.setSeconds(59);

		Date[] dates = new Date[2];
		dates[0] = dayFrom;
		dates[1] = dayTo;
		return dates;
	}

	public static String dateToStrLong(Date date) {
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String dateString = formatter.format(date);
		return dateString;
	}

	public static String dateToStrShort(Date date) {
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
		String dateString = formatter.format(date);
		return dateString;
	}
	
	
    //根据日期计算所在周的上下界 
    public static Map<String, Date> convertWeekByDate(Date time) throws ParseException {  
        Map<String, Date> map = new HashMap<String, Date>();  
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd"); // 设置时间格式  
        Calendar cal = Calendar.getInstance();  
        cal.setTime(time);  
        // 判断要计算的日期是否是周日，如果是则减一天计算周六的，否则会出问题，计算到下一周去了  
        int dayWeek = cal.get(Calendar.DAY_OF_WEEK);// 获得当前日期是一个星期的第几天  
        if (1 == dayWeek) {  
            cal.add(Calendar.DAY_OF_MONTH, -1);  
        }  
        cal.setFirstDayOfWeek(Calendar.MONDAY);// 设置一个星期的第一天，按中国的习惯一个星期的第一天是星期一  
        int day = cal.get(Calendar.DAY_OF_WEEK);// 获得当前日期是一个星期的第几天  
        cal.add(Calendar.DATE, cal.getFirstDayOfWeek() - day);// 根据日历的规则，给当前日期减去星期几与一个星期第一天的差值  
        String imptimeBegin = sdf.format(cal.getTime());  
        // System.out.println("所在周星期一的日期：" + imptimeBegin);  
        cal.add(Calendar.DATE, 6);  
        String imptimeEnd = sdf.format(cal.getTime());  
        // System.out.println("所在周星期日的日期：" + imptimeEnd);  
  
        map.put("first",sdf.parse(imptimeBegin));  
  
        map.put("last", sdf.parse(imptimeEnd));  
  
        return map;  
    }  
  
    //根据日期计算当月的第一天与最后一天 
    public static Map<String, Date> convertMonthByDate(Date date) throws ParseException {  
        Map<String, Date> map = new HashMap<String, Date>();  
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");  
        Calendar calendar = Calendar.getInstance();  
        calendar.setTime(date);  
        calendar.add(Calendar.MONTH, -1);  
        Date theDate = calendar.getTime();  
        // 上个月第一天  
        GregorianCalendar gcLast = (GregorianCalendar) Calendar.getInstance();  
        gcLast.setTime(theDate);  
        gcLast.set(Calendar.DAY_OF_MONTH, 1);  
        String day_first = sdf.format(gcLast.getTime());  
        // 上个月最后一天  
        calendar.add(Calendar.MONTH, 1); // 加一个月  
        calendar.set(Calendar.DATE, 1); // 设置为该月第一天  
        calendar.add(Calendar.DATE, -1); // 再减一天即为上个月最后一天  
        String day_last = sdf.format(calendar.getTime());  
        map.put("first", sdf.parse(day_first));  
        map.put("last", sdf.parse(day_last));  
        return map;  
    }  
  
    //根据日期计算当年的第一天与最后一天  
    public static Map<String, Object> convertYearByDate(Date date) {  
        Map<String, Object> map = new HashMap<String, Object>();  
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");  
        Calendar calendar = Calendar.getInstance();  
        int year = date.getYear() + 1900;  
        calendar.clear();  
        calendar.set(Calendar.YEAR, year);  
        Date currYearFirst = calendar.getTime();  
        calendar.set(Calendar.YEAR, year + 1);  
        calendar.add(calendar.DATE, -1);  
        Date lastYearFirst = calendar.getTime();  
        map.put("first", currYearFirst);  
        map.put("last", lastYearFirst);  
        return map;  
  
    }  
  
	//构建指定时间所在周的map
	public static LinkedHashMap<String,Integer> genWeekSegmentMap(Date date) throws ParseException{
		LinkedHashMap<String, Integer> hashMap = new LinkedHashMap<String,Integer>();
		date= TimeHelper.convertWeekByDate(date).get("first");
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(date);
		for (int i = 0; i < 7; i++) {
			hashMap.put(TimeHelper.dateToStrShort(calendar.getTime()),0);
			calendar.add(Calendar.DATE, 1);
		}
		return hashMap;
	}
	
	
	//构建指定时间所在周的map
	public static LinkedHashMap<String,Integer> genPeriodSegmentMap(Date start,Date end) throws ParseException{
		LinkedHashMap<String, Integer> hashMap = new LinkedHashMap<String,Integer>();
		Integer days=(int)((end.getTime()-start.getTime())/1000/3600/24);
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(start);
		for (int i = 0; i <=days; i++) {
			hashMap.put(TimeHelper.dateToStrShort(calendar.getTime()),0);
			calendar.add(Calendar.DATE, 1);
		}
		return hashMap;
	}
	
	
	
	
	/*@SuppressWarnings("deprecation")
	public static void main(String[] args) throws ParseException {
//		Map<String, Object> weekMap = convertWeekByDate(new Date());
//		Map<String, Object> weekMap = convertMonthByDate(new Date());
//		Map<String, Object> weekMap = convertYearByDate(new Date());
//		
//		Set<Entry<String, Object>> entrySet = weekMap.entrySet();
//		Iterator<Entry<String, Object>> iterator = entrySet.iterator();
//		while (iterator.hasNext()) {
//			Entry<String, Object> next = iterator.next();
//			System.out.println(next.getKey()+":"+next.getValue());
//		}
//		Date start = new Date();
//		start.setMonth(5);
//		LinkedHashMap<String, Integer> genWeekSegmentMap = genPeriodSegmentMap(start,new Date());
//		System.out.println(genWeekSegmentMap);
	}
	*/
	
	
	
	
	
	
}
