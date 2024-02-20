package com.antonioalejandro.utils.excel;

import java.util.TreeSet;
import java.util.stream.Stream;

import com.antonioalejandro.utils.excel.annotations.ExcelColumn;
import com.antonioalejandro.utils.excel.annotations.ExcelItem;
import com.antonioalejandro.utils.excel.interfaces.ExcelCreatorFunction;

public class ExcelCreator {

    public static <T> ExcelCreatorFunction<T> from(Class<T> clazz) {
        if (!clazz.isAnnotationPresent(ExcelItem.class)) {
            return (Iterable<T> iterable) -> new byte[0];
        }

        var fields = clazz.getDeclaredFields();

        if (Stream.of(fields).noneMatch(field -> field.isAnnotationPresent(ExcelColumn.class))) {
            return (Iterable<T> iterable) -> new byte[0];
        }

        var metadataFields = new TreeSet<ExcelData>(
                Stream.of(fields).filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                        .map(field -> new ExcelData(field.getType(), field.getName(),
                                field.getAnnotation(ExcelColumn.class)))
                        .toList());

        final var excelBook = new ExcelBook<T>(clazz.getAnnotation(ExcelItem.class), metadataFields, clazz);

        return excelBook::create;

    }

}
