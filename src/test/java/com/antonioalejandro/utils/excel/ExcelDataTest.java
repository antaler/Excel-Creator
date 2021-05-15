package com.antonioalejandro.utils.excel;

import static org.junit.Assert.*;

import java.util.Date;

import org.junit.Test;

import com.antonioalejandro.utils.excel.enums.ExcelDataTypes;

public class ExcelDataTest {

	@Test
	public void testStringData() throws Exception {
		final var data = "String";
		final var excelData = new ExcelData(data);
		assertEquals(ExcelDataTypes.STRING, excelData.getType());
		assertTrue(excelData.getValue() instanceof String);
		assertEquals(data, excelData.getValue());

	}
	@Test
	public void testLongData() throws Exception {
		final var data = 1L;
		final var excelData = new ExcelData(data);
		assertEquals(ExcelDataTypes.LONG, excelData.getType());
		assertTrue(excelData.getValue() instanceof Long);
		assertEquals(data, excelData.getValue());
	}
	
	@Test
	public void testIntegerData() throws Exception {
		final var data = 1;
		final var excelData = new ExcelData(data);
		assertEquals(ExcelDataTypes.INTEGER, excelData.getType());
		assertTrue(excelData.getValue() instanceof Integer);
		assertEquals(data, excelData.getValue());
	}
	
	@Test
	public void testDoubleData() throws Exception {
		final var data = 1.093d;
		final var excelData = new ExcelData(data);
		assertEquals(ExcelDataTypes.DOUBLE, excelData.getType());
		assertTrue(excelData.getValue() instanceof Double);
		assertEquals(data, excelData.getValue());
	}
	
	@Test
	public void testDateData() throws Exception {
		final var data = new Date();
		final var excelData = new ExcelData(data);
		assertEquals(ExcelDataTypes.DATE, excelData.getType());
		assertTrue(excelData.getValue() instanceof Date);
		assertEquals(data, excelData.getValue());
	}
	
	@Test
	public void testBooleanData() throws Exception {
		final var data = new ExcelBooleanData(false);
		final var excelData = new ExcelData(data);
		assertEquals(ExcelDataTypes.BOOLEAN, excelData.getType());
		assertTrue(excelData.getValue() instanceof ExcelBooleanData);
		assertEquals(data, excelData.getValue());
	}
}
