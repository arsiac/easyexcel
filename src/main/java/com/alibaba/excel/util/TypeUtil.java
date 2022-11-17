package com.alibaba.excel.util;

import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

/**
 * 类型转换工具类
 *
 * @author jipengfei
 */
public class TypeUtil {

    /**
     * 默认支持的日期类型
     */
    private static final List<SimpleDateFormat> DATE_FORMAT_LIST = new LinkedList<>();

    static {
        DATE_FORMAT_LIST.add(createStrict("yyyy/MM/dd HH:mm:ss"));
        DATE_FORMAT_LIST.add(createStrict("yyyy-MM-dd HH:mm:ss"));
        DATE_FORMAT_LIST.add(createStrict("yyyy/MM/dd"));
        DATE_FORMAT_LIST.add(createStrict("yyyy-MM-dd"));
    }

    public static Object convert(String value, Field field, String format, boolean us) {
        if (isEmpty(value)) {
            return null;
        }

        // 字符串
        if (String.class.equals(field.getType())) {
            return value;
        }

        // 时间
        if (Date.class.equals(field.getType())) {
            if (value.contains("-") || value.contains("/") || value.contains(":")) {
                return getSimpleDateFormatDate(value, format);
            } else {
                double d = Double.parseDouble(value);
                return DateUtil.getJavaDate(d, us);
            }
        }

        // 布尔值
        if (Boolean.class.equals(field.getType()) || boolean.class.equals(field.getType())) {
            String valueLower = value.toLowerCase();
            if (valueLower.equals("true") || valueLower.equals("false")) {
                return Boolean.parseBoolean(value.toLowerCase());
            }
            int integer = Integer.parseInt(value);
            return integer != 0;
        }

        return tryNumber(value, field);
    }

    private static Object tryNumber(String value, Field field) {
        if (Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
            return Integer.parseInt(value);
        }
        if (Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
            return Double.parseDouble(value);
        }
        if (Long.class.equals(field.getType()) || long.class.equals(field.getType())) {
            return Long.parseLong(value);
        }
        if (BigDecimal.class.equals(field.getType())) {
            return new BigDecimal(value);
        }
        return null;
    }

    private static Date getSimpleDateFormatDate(String value, String format) {
        if (isEmpty(value)) {
            return null;
        }

        Date date = null;
        if (isEmpty(format)) {
            for (SimpleDateFormat dateFormat : DATE_FORMAT_LIST) {
                try {
                    date = dateFormat.parse(value);
                } catch (ParseException ignored) {
                    // ignore
                }
                if (date != null) {
                    break;
                }
            }
        } else {
            SimpleDateFormat simpleDateFormat = createStrict(format);
            try {
                date = simpleDateFormat.parse(value);
            } catch (ParseException e) {
                throw new IllegalStateException(e);
            }
        }

        return date;
    }

    /**
     * 字符串是否为空
     *
     * @param value 字符串
     * @return true 空字符串; 否则为非空字符串
     */
    private static boolean isEmpty(String value) {
        if (value == null) {
            return true;
        }

        return value.trim().length() == 0;
    }

    /**
     * 创建严格模式的格式化处理
     *
     * @param pattern 时间格式
     * @return 时间处理工具
     */
    private static SimpleDateFormat createStrict(String pattern) {
        final SimpleDateFormat dateFormat = new SimpleDateFormat(pattern);
        dateFormat.setLenient(false);
        return dateFormat;
    }

    private TypeUtil() {
    }

}
