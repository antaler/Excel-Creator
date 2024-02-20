package com.antaler.utils.excel.styles;

import java.awt.Color;
import java.util.Objects;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Styles {

	private Styles() {

	}

	public static XSSFCellStyle header(final XSSFWorkbook workbook, String color) {
		final XSSFCellStyle headerStyle = workbook.createCellStyle();

		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		headerStyle.setBorderTop(BorderStyle.THIN);
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setBorderLeft(BorderStyle.THIN);
		headerStyle.setBorderRight(BorderStyle.THIN);

		headerStyle.setFont(Fonts.header(workbook));
		if (Objects.nonNull(color) && !color.isBlank()) {

			final IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();
			headerStyle.setFillForegroundColor(new XSSFColor(parseColor(color), colorMap));
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		return headerStyle;
	}

	public static XSSFCellStyle data(final XSSFWorkbook workbook, String color) {
		final XSSFCellStyle dataStyle = workbook.createCellStyle();

		dataStyle.setAlignment(HorizontalAlignment.CENTER);

		dataStyle.setBorderTop(BorderStyle.THIN);
		dataStyle.setBorderBottom(BorderStyle.THIN);
		dataStyle.setBorderLeft(BorderStyle.THIN);
		dataStyle.setBorderRight(BorderStyle.THIN);

		if (Objects.nonNull(color) && !color.isBlank()) {

			final IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();
			dataStyle.setFillForegroundColor(new XSSFColor(parseColor(color), colorMap));
			dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}

		return dataStyle;
	}

	public static XSSFCellStyle date(final XSSFWorkbook workbook, String format, String color) {
		final XSSFCellStyle dataStyle = workbook.createCellStyle();

		dataStyle.setAlignment(HorizontalAlignment.CENTER);

		dataStyle.setBorderTop(BorderStyle.THIN);
		dataStyle.setBorderBottom(BorderStyle.THIN);
		dataStyle.setBorderLeft(BorderStyle.THIN);
		dataStyle.setBorderRight(BorderStyle.THIN);
		final XSSFDataFormat df = workbook.createDataFormat();
		dataStyle.setDataFormat(df.getFormat(format));
		if (Objects.nonNull(color) && !color.isBlank()) {

			final IndexedColorMap colorMap = workbook.getStylesSource().getIndexedColors();
			dataStyle.setFillForegroundColor(new XSSFColor(parseColor(color), colorMap));
			dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		return dataStyle;
	}

	private static Color parseColor(String color) {
		var rgb = color.split("_");
		return new Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2]));
	}
}
