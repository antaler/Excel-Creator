package com.antaler.utils.excel;

import com.antaler.utils.excel.annotations.ExcelColumn;

public record ExcelData(Class<?> type, String fieldName, ExcelColumn columnData) implements Comparable<ExcelData> {

	@Override
	public int compareTo(ExcelData other) {
		return Integer.compare(columnData.order(), other.columnData.order());
	}

}
