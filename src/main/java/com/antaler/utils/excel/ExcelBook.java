package com.antaler.utils.excel;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.util.List;
import java.util.SortedSet;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.antaler.utils.excel.annotations.ExcelColumn;
import com.antaler.utils.excel.annotations.ExcelItem;
import com.antaler.utils.excel.styles.Styles;

public class ExcelBook<T> {

	private final XSSFWorkbook book;
	private final XSSFSheet sheet;

	private int headerRow;
	private int headerColumn;
	private int logoRow;
	private int logoColumn;

	private int imageAnchor;
	private int imageAnchor2;

	private int columnsSize;

	private final XSSFCellStyle styleHeader;
	private final XSSFCellStyle styleData;

	private SortedSet<ExcelData> metadata;
	private ExcelItem excelItem;

	public ExcelBook(ExcelItem excelItem, SortedSet<ExcelData> metadata, Class<T> clazz) {
		this.book = new XSSFWorkbook();
		this.sheet = book.createSheet(excelItem.name());
		this.styleHeader = Styles.header(book, excelItem.headerColor());
		this.styleData = Styles.data(book, excelItem.dataColor());
		this.metadata = metadata;
		this.excelItem = excelItem;
		sheet.setDisplayGridlines(excelItem.blank());
		padding();
		addLogo(clazz);
		createHeaders(metadata.stream().map(ExcelData::getColumnData).map(ExcelColumn::title).toList());
	}

	final byte[] create(final Iterable<T> data) {
		data.forEach(this::fillRow);
		autosize();
		try (final var bos = new ByteArrayOutputStream();) {
			book.write(bos);
			return bos.toByteArray();
		} catch (IOException ignored) {
			ignored.printStackTrace();
		} finally {
			try {
				book.close();
			} catch (IOException e) {
				
			}
		}
		return new byte[0];

	}

	private void autosize() {
		sheet.autoSizeColumn(0);
		for (int i = 0; i < columnsSize; i++) {
			sheet.autoSizeColumn(i + headerColumn);
		}
	}

	private byte[] loadFromFile() {
		try (InputStream stream = new FileInputStream(excelItem.logo())) {
			return stream.readAllBytes();
		} catch (Exception e) {
			return new byte[0];
		}
	}

	private byte[] loadFromClasspath(Class<T> clazz) {
		var url = clazz.getClassLoader().getResource(excelItem.classPathLogo());
		try (var inputStream = new FileInputStream(url.getFile())) {
			return inputStream.readAllBytes();
		} catch (Exception e) {
			return new byte[0];
		}
	}

	private void addLogo(Class<T> clazz) {
		byte[] image;

		if (!excelItem.classPathLogo().isBlank()) {
			image = loadFromClasspath(clazz);
		} else if (!excelItem.logo().isBlank()) {
			image = loadFromFile();
		} else {
			return;
		}

		final CreationHelper helper = book.getCreationHelper();
		final int indexLogo = book.addPicture(image, org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_JPEG);
		final Drawing<?> drawing = sheet.createDrawingPatriarch();
		final ClientAnchor anchorLogo = helper.createClientAnchor();

		anchorLogo.setCol1(logoColumn);
		anchorLogo.setRow1(logoRow);
		anchorLogo.setCol2(imageAnchor);
		anchorLogo.setRow2(imageAnchor2);

		anchorLogo.setAnchorType(AnchorType.MOVE_DONT_RESIZE);
		drawing.createPicture(anchorLogo, indexLogo);
		final Picture pict = drawing.createPicture(anchorLogo, indexLogo);
		pict.resize(3, 4);
	}

	private void createHeaders(final List<String> headers) {
		final XSSFRow hr = sheet.createRow(this.headerRow);
		XSSFCell cell;
		this.columnsSize = headers.size();
		for (int i = 0; i < headers.size(); i++) {
			cell = hr.createCell(i + headerColumn);
			cell.setCellValue(headers.get(i));
			cell.setCellStyle(this.styleHeader);
		}
	}

	private XSSFRow nextDataRow() {
		return sheet.createRow(++headerRow);
	}

	private void fillRow(T t) {
		Method method;
		var row = nextDataRow();
		var offset = 0;
		XSSFCell cell;
		for (ExcelData excelData : metadata) {
			cell = row.createCell(headerColumn + offset++);
			method = getMethod(t, excelData);
			addValue(method, cell, excelData, t);
			if (method.getReturnType().equals(LocalDateTime.class)) {
				cell.setCellStyle(
						Styles.date(book, excelData.getColumnData().dateFormat().getFormat(), excelItem.dataColor()));
			} else {
				cell.setCellStyle(styleData);
			}
		}
	}

	private Method getMethod(T t, ExcelData excelData) {
		try {
			return t.getClass().getMethod("get%s".formatted(capitalize(excelData.getFieldName())), null);
		} catch (NoSuchMethodException | SecurityException e) {
			throw new IllegalArgumentException(e);
		}
	}

	private void addValue(Method method, XSSFCell cell, ExcelData excelData, T t) {
		try {
			if (method.getReturnType().equals(String.class)) {
				cell.setCellValue((String) method.invoke(t));
			} else if (method.getReturnType().equals(Double.class)) {
				cell.setCellValue((Double) method.invoke(t));
			} else if (method.getReturnType().equals(Long.class)) {
				cell.setCellValue((Long) method.invoke(t));
			} else if (method.getReturnType().equals(Integer.class)) {
				cell.setCellValue((Integer) method.invoke(t));
			} else if (method.getReturnType().equals(LocalDateTime.class)) {
				cell.setCellValue((LocalDateTime) method.invoke(t));
			} else if (method.getReturnType().equals(Boolean.class)) {
				cell.setCellValue(((boolean) method.invoke(t)) ? excelData.getColumnData().trueValue()
						: excelData.getColumnData().falseValue());
			} else {
				cell.setCellValue(method.invoke(t).toString());
			}

		} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
		}
	}

	private String capitalize(String string) {
		return string.substring(0, 1).toUpperCase().concat(string.substring(1));
	}

	private void padding() {
		if (!excelItem.logo().isBlank() || !excelItem.classPathLogo().isBlank()) { // be logo
			this.headerRow = 6;
			this.headerColumn = 3;
			this.logoRow = 1;
		} else {
			this.headerRow = 0;
			this.headerColumn = 0;
			this.logoRow = 0;
		}
		this.logoColumn = 0;
		this.imageAnchor = 0;
		this.imageAnchor2 = 0;
	}

}
