package com.antonioalejandro.utils.excel.enums;

import lombok.AllArgsConstructor;
import lombok.Getter;

@AllArgsConstructor
public enum ExcelDateFormat {
	NUMBER_SHORT_WITHOUT_TIME("dd-MM-yy"),
 	NUMBER_SHORT_WITH_TIME("dd-MM-yy HH:mm:ss"),
	NUMBER_LONG_WITHOUT_TIME("dd-MM-yyyy"),
	NUMBER_LONG_WITH_TIME("dd-MM-yyyy HH:mm:ss"),
	TIME("HH:mm:ss");

	@Getter
	private final String format;

}
