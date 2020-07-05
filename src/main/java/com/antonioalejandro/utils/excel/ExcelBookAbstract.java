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
	 * Obtiene el tipo de formato que pone a los campos que tenga un Date
	 *
	 * @return
	 */
	public ExcelCellDateFormat getFormatDate() {
		return formatDate;
	}

	/**
	 * Establece el formato de fecha, creando un nuevo estilo de la celda
	 *
	 * @param formatDate
	 */
	public void setFormatDate(final ExcelCellDateFormat formatDate) {
		this.formatDate = formatDate;
		this.styleDate = createDateCellStyle(book);
	}

	/**
	 * Añade el logo en formato byte[]. En caso de que se establezcan los logos con
	 * byte[] e ImputStream, se prioriza el del byte[].
	 *
	 * @param logo
	 */
	public void addLogo(final byte[] logo) {
		this.logoBytes = logo;
	}

	/**
	 * Añade el logo en formato InputStream. En caso de que se establezcan los logos
	 * con byte[] e ImputStream, se prioriza el del byte[].
	 *
	 * @param logo
	 */
	public void addLogo(final InputStream logo) {
		this.logo = logo;
	}

	/**
	 * Cierra el excel y no se pueden hacer añadir mas datos
	 *
	 * @throws IOException
	 */
	public void close() throws IOException {
		book.close();
	}

	// protected methods
	/**
	 * Procesa el excel y lo escribe en la ruta pasada como parametro. Si no hay
	 * logo el excel empieza en el A1.
	 *
	 * @param data    Lista de objetos que tengan que implementen la interfaz
	 *                <b>IExcelObject</b>
	 * @param headers Lista de Strings que continene los nombres de las cabeceras.
	 * @param path    Direccion del fichero donde se guaradara el excel. <strong>SIN
	 *                EXTENSION</strong>
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
	 * Procesa el excel para poder enviarlo con un byte[]. Si no hay logo el excel
	 * empieza en el A1.
	 *
	 * @param data    Lista de objetos que tengan que implementen la interfaz
	 *                <b>IExcelObject</b>
	 * @param headers Lista de Strings que continene los nombres de las cabeceras.
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
	 * Establece el color al header
	 *
	 * @param color
	 */
	protected final void setColorHeader(final Color color) {
		final IndexedColorMap colorMap = book.getStylesSource().getIndexedColors();
		this.styleHeader.setFillForegroundColor(new XSSFColor(color, colorMap));
		this.styleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	}

	/**
	 * Establece el color a las celdas de los datos
	 *
	 * @param color
	 */
	protected final void setColorData(final Color color) {
		final IndexedColorMap colorMap = book.getStylesSource().getIndexedColors();
		this.styleData.setFillForegroundColor(new XSSFColor(color, colorMap));
		this.styleData.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	}

	/**
	 * Activa o desactiva el grid
	 *
	 * @param hasGridLines
	 */
	protected void setLinesSheet(final boolean hasGridLines) {
		sheet.setDisplayGridlines(hasGridLines);
	}

	// private methods
	/**
	 * EScribbe el excel en un fichero
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
	 * Ajusta las columnas al tamaño de los datos
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
	 * Inserta el logo en el excel. Se le da prioridad al logo en formato byte[]
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
	 * Crear la fila del header y las columnas con los titulos.
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
	 * Establece el estilo del header a una celda
	 *
	 * @param cell
	 */
	private void setStyleHeader(final XSSFCell cell) {
		cell.setCellStyle(this.styleHeader);
	}

	/**
	 * Escribe los datos en en excel. NO crea un fichero
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
			objectData = data.get(i).getFields();
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
	 * Crea un numero de columnas pasado como parametro
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
	 * Crear el estilo de las celdas que contengan fechas
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
	 * Crea el estilo de las celdas que contengan datos
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
	 * Crear el estilo de las celdas que sean el header
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
	 * Crear una fuente con el texto en negrita
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
	 * EStablece los indices para la isnercion de los datos y del logo si lo hubiese
	 * al ahora de ccrear el excel
	 */
	private void setRowsAndColumnsIndex() {
		if (logo != null && logoBytes != null) { // hay logo
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
