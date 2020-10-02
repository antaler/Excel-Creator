package com.antonioalejandro.utils.excel;

import java.awt.Color;
import java.io.IOException;
import java.util.List;

import com.antonioalejandro.utils.excel.interfaces.IExcelObject;


/**

 * This class is a implementation of the ExcelBookAbstract class

 * @author: Antonio Alejandro Serrano Ramírez

 * @version: 1.0

 * @see <a href = "http://www.antonioalejandro.com" /> www.antonioalejandro.com</a>

 */

public final class ExcelBook<T extends IExcelObject> extends ExcelBookAbstract {

	private List<String> headers;
	private List<T> data;

	public ExcelBook(final List<String> headers, final List<T> data, final String sheetName) {
		super(sheetName);
		this.headers = headers;
		this.data = data;
	}

	public ExcelBook(final List<String> headers, final String sheetName) {
		super(sheetName);
		this.headers = headers;
	}

	public ExcelBook(final String sheetName) {
		super(sheetName);
	}

	/**
	 * Set the headers that the excel will have
	 *
	 * @param headers
	 */
	public void setHeaders(final List<String> headers) {
		this.headers = headers;
	}

	/**
	 * Set the data
	 *
	 * @param data List with object that implement interface
	 */
	public void setData(final List<T> data) {
		this.data = data;
	}
	/**
	 * Get headers that was established
	 */
	@Override
	public List<String> getHeaders() {
		return headers;
	}

	/**
	 * Creates the excel file in the received path as a parameter.
	 *
	 * @param path
	 * @throws IOException
	 */
	public void write(final String path) throws IOException {
		processBook(data, headers, path);
	}

	/**
	 * Return excel in a <b>byte[]</b>.
	 *
	 * @return byte[]
	 * @throws IOException
	 */
	public byte[] prepareToSend() throws IOException {
		return processBookToSend(data, headers);
	}

	/**
	 * Sets the background color of the header
	 *
	 * @param color
	 */
	public void setHeaderColor(final Color color) {
		setColorHeader(color);
	}

	/**
	 * Sets the background color of the data cells
	 *
	 * @param color
	 */
	public void setDataColor(final Color color) {
		setColorData(color);
	}

	/**
	 * Remove all the lines in the excel except those that are filled in with data (the header and data cells).
	 */
	public void setBlankSheet() {
		setLinesSheet(false);
	}

}
