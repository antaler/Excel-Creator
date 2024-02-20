package com.antaler.utils.excel;

import com.antaler.utils.excel.annotations.ExcelColumn;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class ExcelData implements Comparable<ExcelData> {

	private Class<?> type;

	private String fieldName;

	private ExcelColumn columnData;


	@Override
	public int compareTo(ExcelData other) {
		return Integer.compare(columnData.order(), other.columnData.order());
	}

}
