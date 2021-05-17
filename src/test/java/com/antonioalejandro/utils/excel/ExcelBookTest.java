package com.antonioalejandro.utils.excel;

import static org.junit.Assert.assertNotNull;

import java.awt.Color;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.junit.Test;

import com.antonioalejandro.utils.excel.interfaces.IExcelObject;

public class ExcelBookTest {

	@Test
	public void testConstructorSheetName() throws Exception {
		class Test implements IExcelObject {
			@Override
			public List<ExcelData> obtainFields() {
				var list = new ArrayList<ExcelData>();
				list.add(new ExcelData("NOMBRE"));
				list.add(new ExcelData(new Date()));
				list.add(new ExcelData(2.3));
				list.add(new ExcelData(1));
				list.add(new ExcelData(12L));
				list.add(new ExcelData(new ExcelBooleanData(true)));
				list.add(new ExcelData(new ExcelBooleanData(false)));
				return list;

			}
		}
		var o = new Test();
		var list = List.of(o);

		var book = new ExcelBook<Test>("NAME");

		book.setBlankSheet();
		book.setData(list);
		book.setHeaders(List.of("h1", "h2", "h3", "h4", "h6","h7"));
		book.setHeaderColor(new Color(1, 2, 4, 0));
		book.setDataColor(new Color(1, 123, 123));

		assertNotNull(book.getHeaders());
		assertNotNull(book.prepareToSend());
		book = new ExcelBook<>(List.of("h1", "h2", "h3", "h4", "h6","h7"), "NAME 2");
		book.setData(list);
		assertNotNull(book.prepareToSend());
		assertNotNull(new ExcelBook<>(List.of("h1", "h2", "h3", "h4", "h6","h7"), list, "NAME 3"));

	}
}
