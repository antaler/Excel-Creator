package com.antonioalejandro.utils.excel.enums;

public enum ExcelCellDateFormat {
	NUMBER_SHORT_WITHOUT_TIME, NUMBER_SHORT_WITH_TIME, NUMBER_LONG_WITHOUT_TIME, NUMBER_LONG_WITH_TIME, TIME;

	/**
	 * Obtiene el strign del formato a partir del valor del propio enum
	 * 
	 * @param format
	 * @return
	 */
	public static String getFormat(final ExcelCellDateFormat format) {
		final String aux;
		switch (format) {
		case NUMBER_SHORT_WITHOUT_TIME:
			aux = "dd-MM-yy";
			break;
		case NUMBER_SHORT_WITH_TIME:
			aux = "dd-MM-yy HH:mm:ss";
			break;
		case NUMBER_LONG_WITHOUT_TIME:
			aux = "dd-MM-yyyy";
			break;
		case NUMBER_LONG_WITH_TIME:
			aux = "dd-MM-yyyy HH:mm:ss";
			break;
		case TIME:
			aux = "HH:mm:ss";
			break;
		default:
			aux = "dd-MM-yy";
			break;
		}
		return aux;
	}

}
