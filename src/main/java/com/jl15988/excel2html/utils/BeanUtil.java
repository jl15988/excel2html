package com.jl15988.excel2html.utils;

import com.jl15988.excel2html.stream.MethodStream;

import java.lang.reflect.Constructor;
import java.lang.reflect.Method;

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

//    /**
//     * 执行指定类的方法获取结果，没有指定方法返回 null
//     *
//     * @param clazz       指定类
//     * @param methodName  方法名称
//     * @param resultClass 结果类型类
//     * @param args        方法参数
//     * @param <T>         指定类
//     * @param <R>         执行方法结果类型
//     * @return 结果
//     */
//    public static <T, R> R exeMethod(Class<T> clazz, String methodName, Class<R> resultClass, Object... args) throws InvocationTargetException, IllegalAccessException {
//        Method[] methods = clazz.getMethods();
//        for (Method method : methods) {
//            method.setAccessible(true);
//            if (method.getName().equals(methodName)) {
//                return (R) method.invoke(clazz);
//            }
//        }
//        return null;
//    }

    /**
     * 执行指定类的方法获取结果，没有指定方法返回 null
     *
     * @param clazz      指定类
     * @param methodName 方法名称
     * @param <T>        指定类
     * @return 方法流转
     */
    public static <T> MethodStream exeMethod(Class<T> clazz, String methodName) {
        Method[] methods = clazz.getMethods();
        for (Method method : methods) {
            method.setAccessible(true);
            if (method.getName().equals(methodName)) {
                return new MethodStream(clazz, method);
            }
        }
        return new MethodStream(clazz, null);
    }
}
