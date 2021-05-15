package com.antonioalejandro.utils.excel;

import static org.junit.Assert.*;

import org.junit.Test;

public class ExcelBooleanDataTest {

	@Test
	public void testDefault() throws Exception {
		var excelBooleanData = new ExcelBooleanData(true);

		assertEquals("True", excelBooleanData.getTextTrueOption());
		assertEquals("False", excelBooleanData.getTextFalseOption());
		assertTrue(excelBooleanData.getBooleanValue());

		excelBooleanData.setBooleanValue(false);

		assertFalse(excelBooleanData.getBooleanValue());

		excelBooleanData.setTextTrueOption("Verdadero");
		excelBooleanData.setTextFalseOption("Falso");

		assertEquals("Verdadero", excelBooleanData.getTextTrueOption());
		assertEquals("Falso", excelBooleanData.getTextFalseOption());

	}

	@Test
	public void testCustomConstructor() throws Exception {
		var excelBooleanData = new ExcelBooleanData(false, "Verdadero", "Falso");
		assertEquals("Verdadero", excelBooleanData.getTextTrueOption());
		assertEquals("Falso", excelBooleanData.getTextFalseOption());
		assertFalse(excelBooleanData.getBooleanValue());
	}
}
