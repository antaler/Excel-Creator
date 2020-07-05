package com.antonioalejandro.utils.excel;

import java.awt.Color;
import java.io.IOException;
import java.util.List;

import com.antonioalejandro.utils.excel.interfaces.IExcelObject;

public class ExcelBook<T extends IExcelObject> extends ExcelBookAbstract {

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
	 * Establece la cabecera que tendrá el excel
	 *
	 * @param headers
	 */
	public void setHeaders(final List<String> headers) {
		this.headers = headers;
	}

	/**
	 * Establece los datos con los que se trabajaran
	 *
	 * @param data
	 */
	public void setData(final List<T> data) {
		this.data = data;
	}

	@Override
	public List<String> getHeaders() {
		return headers;
	}

	/**
	 * Crea el fichero excel en la ruta pasada como parámetro.
	 *
	 * @param path
	 * @throws IOException
	 */
	public void write(final String path) throws IOException {
		processBook(data, headers, path);
	}

	/**
	 * Devuelve el excel en un <b>byte[]</b>.
	 *
	 * @return byte[]
	 * @throws IOException si no se ha coneguido convertir a byte[]
	 */
	public byte[] prepareToSend() throws IOException {
		return processBookToSend(data, headers);
	}

	/**
	 * Establece el color que tiene de fondo la cabecera
	 *
	 * @param color
	 */
	public void setHeaderColor(final Color color) {
		setColorHeader(color);
	}

	/**
	 * Establece el color de las celdas que tiene los datos.No afecta a las de la
	 * cabecera
	 *
	 * @param color
	 */
	public void setDataColor(final Color color) {
		setColorData(color);
	}

	/**
	 * Quita todas las lineas del excel excepto las que esten rellenas con datos(las
	 * celdas del header y las celdas de los datos en sí).
	 */
	public void setBlankSheet() {
		setLinesSheet(false);
	}

}
