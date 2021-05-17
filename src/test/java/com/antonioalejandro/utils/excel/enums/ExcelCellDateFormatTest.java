package com.antonioalejandro.utils.excel.enums;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class ExcelCellDateFormatTest {

	@Test
	public void testGetFormatSHORT() throws Exception {
		assertEquals("dd-MM-yy", ExcelCellDateFormat.getFormat(ExcelCellDateFormat.NUMBER_SHORT_WITHOUT_TIME));
		assertEquals("dd-MM-yy HH:mm:ss", ExcelCellDateFormat.getFormat(ExcelCellDateFormat.NUMBER_SHORT_WITH_TIME));
	}
	
	@Test
	public void testGetFormatLONG() throws Exception {
		assertEquals("dd-MM-yyyy", ExcelCellDateFormat.getFormat(ExcelCellDateFormat.NUMBER_LONG_WITHOUT_TIME));
		assertEquals("dd-MM-yyyy HH:mm:ss", ExcelCellDateFormat.getFormat(ExcelCellDateFormat.NUMBER_LONG_WITH_TIME));
	}
	
	@Test
	public void testGetFormatTIME() throws Exception {
		assertEquals("HH:mm:ss", ExcelCellDateFormat.getFormat(ExcelCellDateFormat.TIME));
	}
}
