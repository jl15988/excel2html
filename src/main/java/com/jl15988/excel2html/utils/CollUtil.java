package com.jl15988.excel2html.utils;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * @author Jalon
 * @since 2024/12/1 12:19
 **/
public class CollUtil {

    public static <T> String join(List<T> list, String separator) {
        if (list == null) return "";
        if (list.isEmpty()) return "";
        StringBuilder resultString = new StringBuilder();
        for (int i = 0; i < list.size(); i++) {
            T item = list.get(i);
            resultString.append(item == null ? "null" : item.toString());
            if (i < list.size() - 1) resultString.append(separator);
        }
        return resultString.toString();
    }

    public static <T> String join(Iterable<T> iterable, CharSequence conjunction) {
        if (iterable == null) return "";

        StringBuilder resultString = new StringBuilder();

        Iterator<T> iterator = iterable.iterator();
        List<String> appendResult = append(iterator);

        for (int i = 0; i < appendResult.size(); i++) {
            String item = appendResult.get(i);
            resultString.append(item);
            if (i < appendResult.size() - 1) {
                resultString.append(conjunction);
            }
        }
        return resultString.toString();
    }

    public static <T> List<String> append(Iterator<T> iterator) {
        List<String> resultList = new ArrayList<String>();
        if (null != iterator) {
            while (iterator.hasNext()) {
                T obj = iterator.next();
                if (null == obj) {
                    resultList.add("null");
                } else if (obj.getClass().isArray()) {
                    List<String> appendResult = append(Arrays.asList(obj).iterator());
                    resultList.addAll(appendResult);
                } else if (obj instanceof Iterator) {
                    List<String> appendResult = append((Iterator) obj);
                    resultList.addAll(appendResult);
                } else if (obj instanceof Iterable) {
                    List<String> appendResult = append(((Iterable) obj).iterator());
                    resultList.addAll(appendResult);
                } else {
                    resultList.add(obj.toString());
                }
            }
        }
        return resultList;
    }
}
