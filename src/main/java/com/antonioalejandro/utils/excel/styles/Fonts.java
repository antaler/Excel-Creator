package com.antonioalejandro.utils.excel.styles;

import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Fonts {

	protected static XSSFFont header(final XSSFWorkbook book) {
		final XSSFFont font = book.createFont();
		font.setBold(true);
		return font;
	}
}
