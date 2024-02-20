package com.antonioalejandro.utils.excel.interfaces;

@FunctionalInterface
public interface ExcelCreatorFunction<T> {

	byte[] create(Iterable<T> items);

}
