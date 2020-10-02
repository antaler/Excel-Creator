package com.antonioalejandro.utils.excel;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.antonioalejandro.utils.excel.enums.ExcelCellDateFormat;
import com.antonioalejandro.utils.excel.interfaces.IExcelObject;

/**
 * 
 * This class is a abstraction of the excel book
 * 
 * @author: Antonio Alejandro Serrano Ramírez
 * 
 * @version: 1.0
 * 
 * @see <a href = "http://www.antonioalejandro.com" />
 *      www.antonioalejandro.com</a>
 * 
 */

public abstract class ExcelBookAbstract<T extends IExcelObject> {

	private final XSSFWorkbook book;
	private final XSSFSheet sheet;

	private int headerRow;
	private int headerColumn;
	private int logoRow;
	private int logoColumn;

	private int imageAnchor;
	private int imageAnchor2;

	private final XSSFCellStyle styleHeader;
	private final XSSFCellStyle styleData;
	private XSSFCellStyle styleDate;

	private ExcelCellDateFormat formatDate;

	private InputStream logo;
	private byte[] logoBytes;

	public static final String EXCEL_EXTENSION = ".xlsx";

	public ExcelBookAbstract(final String sheetName) {
		this.book = new XSSFWorkbook();
		this.sheet = book.createSheet(sheetName);
		this.styleHeader = createHeaderCellStyle(book, createHeaderFont(book));
		this.styleData = createDataCellStyle(book);
	}

	// public methods

	public abstract List<String> getHeaders();

	/**
	 * It gets the type of formatting that puts the fields that have a Date
	 *
	 * @return ExcelCellDatFormat
	 */
	public ExcelCellDateFormat getFormatDate() {
		return formatDate;
	}

	/**
	 * Sets the date format, creating a new style of the cell
	 *
	 * @param formatDate
	 */
	public void setFormatDate(final ExcelCellDateFormat formatDate) {
		this.formatDate = formatDate;
		this.styleDate = createDateCellStyle(book);
	}

	/**
	 * Add the logo in byte[] format. If the logo is set with byte[] and
	 * ImputStream, the byte[] logo is prioritized.
	 *
	 * @param logo
	 */
	public void addLogo(final byte[] logo) {
		this.logoBytes = logo;
	}

	/**
	 * 
	 * Add the logo in InputStream format. If the logo is set with byte[] and
	 * ImputStream, the byte[] logo is prioritized.
	 *
	 * @param logo
	 */
	public void addLogo(final InputStream logo) {
		this.logo = logo;
	}

	/**
	 * Close Excel and no more data can be added
	 *
	 * @throws IOException
	 */
	public void close() throws IOException {
		book.close();
	}

	// protected methods

	/**
	 * Processes the excel and writes it in the last path as a parameter. If there
	 * is no logo the excel starts at A1.
	 *
	 * @param data    List of objects that have the <b>IExcelObject</b> interface
	 *                implemented
	 * @param headers String list containing the names of the headers.
	 * @param path    Address of the file where the excel will be saved
	 *                <strong>without extension</strong>
	 * @throws IOException
	 */
	protected final void processBook(final List<T> data, final List<String> headers, final String path)
			throws IOException {
		setRowsAndColumnsIndex();
		if (logo != null && logoBytes != null) {
			setLogo();
		}
		createHeaders(headers);
		writeData(data);
		autosize(headers.size());
		write(path + EXCEL_EXTENSION);
	}

	/**
	 * It processes the excel to be able to send it with a byte . If there is no logo the excel starts at A1.
	 * 
	 * @param data    List of objects that have the <b>IExcelObject</b> interface
	 *                implemented
	 * @param headers String list containing the names of the headers.
	 *
	 * @return byte[]
	 *
	 * @throws IOException
	 */
	protected final byte[] processBookToSend(final List<T> data, final List<String> headers) throws IOException {
		setRowsAndColumnsIndex();
		if (logo != null && logoBytes != null) {
			setLogo();
		}
		createHeaders(headers);
		writeData(data);
		autosize(headers.size());
		final ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try {
			book.write(bos);
		} finally {
			bos.close();
		}
		return bos.toByteArray();
	}

	/**
	 * Set header row color
	 *
	 * @param color
	 */
	protected final void setColorHeader(final Color color) {
		final IndexedColorMap colorMap = book.getStylesSource().getIndexedColors();
		this.styleHeader.setFillForegroundColor(new XSSFColor(color, colorMap));
		this.styleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	}

	/**
	 * Set data cells color
	 *
	 * @param color
	 */
	protected final void setColorData(final Color color) {
		final IndexedColorMap colorMap = book.getStylesSource().getIndexedColors();
		this.styleData.setFillForegroundColor(new XSSFColor(color, colorMap));
		this.styleData.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	}

	/**
	 * Activate or deactivate the grid
	 *
	 * @param hasGridLines
	 */
	protected void setLinesSheet(final boolean hasGridLines) {
		sheet.setDisplayGridlines(hasGridLines);
	}

	// private methods
	
	/**
	 * Write excel in path
	 *
	 * @param path
	 * @throws IOException
	 */
	private void write(final String path) throws IOException {
		final FileOutputStream fileOutputStream = new FileOutputStream(path);
		this.book.write(fileOutputStream);
		fileOutputStream.close();
	}

	/**
	 * Adjusts columns to data size
	 *
	 * @param num
	 */
	private void autosize(final int num) {
		for (int i = 0; i < num; i++) {
			sheet.autoSizeColumn(i + headerColumn);
		}
		sheet.autoSizeColumn(0);
	}

	/**
	 * Insert the logo in the excel. Priority byte format .
	 *
	 * @throws IOException
	 */
	private void setLogo() throws IOException {
		byte[] imagen;
		if (logoBytes != null) {
			imagen = logoBytes;
		} else {
			imagen = IOUtils.toByteArray(logo);
		}

		final CreationHelper helper = book.getCreationHelper();
		final int indexLogo = book.addPicture(imagen, HSSFWorkbook.PICTURE_TYPE_JPEG);
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

	/**
	 * Create the header row and columns with the titles.
	 *
	 * @param headers
	 */
	private void createHeaders(final List<String> headers) {
		final XSSFRow headerRow = sheet.createRow(this.headerRow);
		XSSFCell cell;
		for (int i = 0; i < headers.size(); i++) {
			cell = headerRow.createCell(i + headerColumn);
			cell.setCellValue(headers.get(i));
			setStyleHeader(cell);
		}
	}

	/**
	 * Set the style of the header to single cell
	 *
	 * @param cell
	 */
	private void setStyleHeader(final XSSFCell cell) {
		cell.setCellStyle(this.styleHeader);
	}

	/**
	 * Write the data in excel. It does <b>NOT</b> create a file
	 *
	 * @param data
	 */
	private void writeData(final List<T> data) {
		final List<XSSFRow> rows = createRowData(data.size());
		List<ExcelData> objectData = null;
		ExcelData excelDataCell;
		XSSFCell cell;
		ExcelBooleanData booleanData;
		for (int i = 0; i < rows.size(); i++) {
			objectData = data.get(i).obtainFields();
			for (int j = 0; j < objectData.size(); j++) {
				excelDataCell = objectData.get(j);
				cell = rows.get(i).createCell(j + headerColumn);
				switch (excelDataCell.getType()) {
				case DOUBLE:
					cell.setCellValue((Double) excelDataCell.getValue());
					cell.setCellStyle(styleData);
					break;
				case STRING:
					cell.setCellValue((String) excelDataCell.getValue());
					cell.setCellStyle(styleData);
					break;
				case LONG:
					cell.setCellValue((Long) excelDataCell.getValue());
					cell.setCellStyle(styleData);
					break;
				case INTEGER:
					cell.setCellValue((Integer) excelDataCell.getValue());
					cell.setCellStyle(styleData);
					break;
				case DATE:
					cell.setCellValue((Date) excelDataCell.getValue());
					cell.setCellStyle(styleDate);
					break;
				case BOOLEAN:
					booleanData = (ExcelBooleanData) excelDataCell.getValue();
					cell.setCellValue(booleanData.getBooleanValue() ? booleanData.getTextTrueOption()
							: booleanData.getTextFalseOption());
					cell.setCellStyle(styleData);
					break;
				default:
					System.err.println("Type data incompatible");
					break;
				}

			}

		}
	}

	/**
	 * Creates a number of received columns as a parameter
	 *
	 * @param num
	 * @return
	 */
	private List<XSSFRow> createRowData(final int num) {
		final ArrayList<XSSFRow> rows = new ArrayList<>(num);
		for (int i = 0; i < num; i++) {
			rows.add(sheet.createRow(i + headerRow + 1));
		}
		return rows;

	}

	/**
	 * Create the style of cells containing dates
	 *
	 * @param workbook
	 * @return
	 */
	private XSSFCellStyle createDateCellStyle(final XSSFWorkbook workbook) {
		final XSSFCellStyle dataStyle = workbook.createCellStyle();

		dataStyle.setAlignment(HorizontalAlignment.CENTER);

		dataStyle.setBorderTop(BorderStyle.THIN);
		dataStyle.setBorderBottom(BorderStyle.THIN);
		dataStyle.setBorderLeft(BorderStyle.THIN);
		dataStyle.setBorderRight(BorderStyle.THIN);
		final XSSFDataFormat df = workbook.createDataFormat();
		dataStyle.setDataFormat(df.getFormat(ExcelCellDateFormat.getFormat(formatDate)));
		return dataStyle;
	}

	/**
	 * Create the style of the cells containing data
	 *
	 * @param workbook
	 * @return
	 */
	private XSSFCellStyle createDataCellStyle(final XSSFWorkbook workbook) {
		final XSSFCellStyle dataStyle = workbook.createCellStyle();

		dataStyle.setAlignment(HorizontalAlignment.CENTER);

		dataStyle.setBorderTop(BorderStyle.THIN);
		dataStyle.setBorderBottom(BorderStyle.THIN);
		dataStyle.setBorderLeft(BorderStyle.THIN);
		dataStyle.setBorderRight(BorderStyle.THIN);

		return dataStyle;
	}

	/**
	 * Create the style of the cells that are in the header
	 *
	 * @param workbook
	 * @param font
	 * @return
	 */
	private XSSFCellStyle createHeaderCellStyle(final XSSFWorkbook workbook, final XSSFFont font) {
		final XSSFCellStyle headerStyle = workbook.createCellStyle();

		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		headerStyle.setBorderTop(BorderStyle.THIN);
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setBorderLeft(BorderStyle.THIN);
		headerStyle.setBorderRight(BorderStyle.THIN);

		if (font != null) {
			headerStyle.setFont(font);
		}

		return headerStyle;
	}

	/**
	 * Create a font with bold text
	 *
	 * @param book
	 * @return
	 */
	private XSSFFont createHeaderFont(final XSSFWorkbook book) {
		final XSSFFont font = book.createFont();
		font.setBold(true);
		return font;
	}

	/**
	 * Establishes the indexes for the insertion of the data and the logo, if it exists, when creating the excel
	 */
	private void setRowsAndColumnsIndex() {
		if (logo != null && logoBytes != null) { // be logo
			this.headerRow = 6;
			this.headerColumn = 3;
			this.logoRow = 1;
			this.logoColumn = 0;
			this.imageAnchor = 0;
			this.imageAnchor2 = 0;
		} else {
			this.headerRow = 0;
			this.headerColumn = 0;
			this.logoRow = 0;
			this.logoColumn = 0;
			this.imageAnchor = 0;
			this.imageAnchor2 = 0;
		}
	}

}
