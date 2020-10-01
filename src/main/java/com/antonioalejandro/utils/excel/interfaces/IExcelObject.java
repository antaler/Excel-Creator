package com.antonioalejandro.utils.excel.interfaces;

import java.util.List;

import com.antonioalejandro.utils.excel.ExcelData;

public interface IExcelObject {
	/**
	 * Gets the values of the object's fields in an ExcelData. The order is respected.
	 * @return fields
	 */
	public List<ExcelData> obtainFields();

}
