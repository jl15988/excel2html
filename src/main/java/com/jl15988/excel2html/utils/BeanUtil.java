package com.jl15988.excel2html.utils;

import java.lang.reflect.Constructor;

/**
 * @author Jalon
 * @since 2024/12/25 13:32
 **/
public class BeanUtil {

    /**
     * 根据目标类创建目标实例
     *
     * @param targetClass    目标类
     * @param parameterTypes 构造参数类型数组
     * @param initargs       构造函数参数值数组
     * @param <T>            目标
     * @return 目标实例
     */
    public static <T> T newInstance(Class<T> targetClass, Class<?>[] parameterTypes, Object[] initargs) {
        try {
            Constructor<T> constructor = targetClass.getConstructor(parameterTypes);
            return constructor.newInstance(initargs);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 根据目标类调用无参构造函数创建目标实例
     *
     * @param targetClass 目标类
     * @param <T>         目标
     * @return 目标实例
     */
    public static <T> T newInstance(Class<T> targetClass) {
        return BeanUtil.newInstance(targetClass, new Class[]{}, new Object[]{});
    }
}
