package com.excel.utils;

import java.lang.reflect.InvocationTargetException;
import java.util.Date;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import javax.servlet.http.HttpServletRequest;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.beanutils.converters.DateConverter;
import org.apache.commons.beanutils.expression.DefaultResolver;

/**
 * Created by bonismo@hotmail.com
 * 上午10:32 on 17/11/1.
 * <p>
 * Bean辅助工具，可以从Form表单中直接读取bean对象，也可以为bean对象设置多级属性
 */
public class BeanHelper {

    /**
     * 处理日期格式
     */
    static {
        DateConverter converter = new DateConverter();
        converter.setPatterns(new String[]{"yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd"});
        ConvertUtils.register(converter, Date.class);
    }

    /**
     * 使用表单数据构建bean对象（只注入名称匹配的属性）
     *
     * @param request
     * @param c
     * @return
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws InstantiationException
     * @throws SecurityException
     * @throws NoSuchMethodException
     */
    public static <T> T bean(HttpServletRequest request, Class<T> c) throws IllegalAccessException, InvocationTargetException, InstantiationException, SecurityException, NoSuchMethodException {
        return (T) bean(request, null, c);
    }

    /**
     * 使用表单数据构建bean对象（只注入名称匹配的属性）
     * prefix用于指定前缀,当值为a时会处理以a.开头的参数（例如a.name,a.age），null或""时代表无前缀
     *
     * @param request
     * @param prefix
     * @param c
     * @return
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws InstantiationException
     * @throws SecurityException
     * @throws NoSuchMethodException
     */
    public static <T> T bean(HttpServletRequest request, String prefix, Class<T> c) throws IllegalAccessException, InvocationTargetException, InstantiationException, SecurityException, NoSuchMethodException {
        T obj = c.newInstance();
        Enumeration<?> pNames = request.getParameterNames();
        while (pNames.hasMoreElements()) {
            String name = (String) pNames.nextElement();
            String[] value = request.getParameterValues(name);
            if (prefix != null && !prefix.isEmpty()) {
                if (name.startsWith(prefix + ".")) {
                    name = name.substring(prefix.length() + 1);
                } else {
                    break;
                }
            }
            setProperty(obj, name, value);
        }
        DefaultResolver rs = new DefaultResolver();

        return (T) obj;
    }

    /**
     * 设置对象的属性,pname支持多级，value可以是数组或对象
     * 例1：setProperty(obj, "username","悟空");
     * 例2：setProperty(obj, "category.parent.id",12);
     * 例3：setProperty(obj, "title",12);
     *
     * @param obj
     * @param pname
     * @param value
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws InstantiationException
     * @throws NoSuchMethodException
     */
    static void setProperty(Object obj, String pname, Object value) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException, InstantiationException {
        if (pname.contains(".")) {
            String fn1 = pname.substring(0, pname.indexOf("."));
            String fn2 = pname.substring(pname.indexOf(".") + 1, pname.length());
            if (PropertyUtils.getPropertyType(obj, fn1) == null) return;
            BeanUtils.setProperty(obj, fn1, PropertyUtils.getPropertyType(obj, fn1).newInstance());
            setProperty(PropertyUtils.getProperty(obj, fn1), fn2, value);
        } else {
            BeanUtils.setProperty(obj, pname, value);
        }
    }

    /**
     * 使用表单数据构建bean对象（只注入名称匹配的属性）
     *
     * @param c
     * @return
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws InstantiationException
     * @throws SecurityException
     * @throws NoSuchMethodException
     */
    public static <T> T bean(Map<String, Object> map, Class<T> c) throws IllegalAccessException, InvocationTargetException, InstantiationException, SecurityException, NoSuchMethodException {
        return bean(map, null, c);
    }

    /**
     * 使用表单数据构建bean对象（只注入名称匹配的属性）
     * prefix用于指定前缀,当值为a时会处理以a.开头的参数（例如a.name,a.age），null或""时代表无前缀
     *
     * @param prefix
     * @param c
     * @return
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     * @throws InstantiationException
     * @throws SecurityException
     * @throws NoSuchMethodException
     */
    public static <T> T bean(Map<String, Object> map, String prefix, Class<T> c) throws IllegalAccessException, InvocationTargetException, InstantiationException, SecurityException, NoSuchMethodException {
        T obj = c.newInstance();
        Iterator<Entry<String, Object>> it = map.entrySet().iterator();
        while (it.hasNext()) {
            Entry<String, Object> en = it.next();
            String name = en.getKey();
            Object value = en.getValue();
            if (prefix != null && !prefix.isEmpty()) {
                if (name.startsWith(prefix + ".")) {
                    name = name.substring(prefix.length() + 1);
                } else {
                    break;
                }
            }
            setProperty(obj, name, value);
        }
        return (T) obj;
    }
}

