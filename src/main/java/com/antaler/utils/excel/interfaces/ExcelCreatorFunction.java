package com.antaler.utils.excel.interfaces;

@FunctionalInterface
public interface ExcelCreatorFunction<T> {

	byte[] create(Iterable<T> items);

}
