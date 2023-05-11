package br.com.haetten.sheetreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import javax.swing.JFileChooser;

import org.apache.poi.hssf.extractor.OldExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
	public static void main(String... args) throws InvalidFormatException {
		JFileChooser fc = new JFileChooser();

		if (fc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
			try {
				readWorkbook(fc.getSelectedFile());
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}

	private static String getCellValue(Cell cell) {
        StringBuilder text = new StringBuilder();
        HSSFDataFormatter _formatter = new HSSFDataFormatter();
        
		switch (cell.getCellType()) {
			case STRING:
				text.append(cell.getRichStringCellValue().getString());
				break;
			case NUMERIC:
				text.append(_formatter.formatCellValue(cell));
				break;
			case BOOLEAN:
				text.append(cell.getBooleanCellValue());
				break;
			case ERROR:
				text.append(ErrorEval.getText(cell.getErrorCellValue()));
				break;
			case FORMULA:
				switch (cell.getCachedFormulaResultType()) {
					case STRING:
						RichTextString str = cell.getRichStringCellValue();
						if (str != null && str.length() > 0) {
							text.append(str);
						}
						break;
					case NUMERIC:
						CellStyle style = cell.getCellStyle();
						double nVal = cell.getNumericCellValue();
						short df = style.getDataFormat();
						String dfs = style.getDataFormatString();
						text.append(_formatter.formatRawCellContents(nVal, df, dfs));
						break;
					case BOOLEAN:
						text.append(cell.getBooleanCellValue());
						break;
					case ERROR:
						text.append(ErrorEval.getText(cell.getErrorCellValue()));
						break;
					default:
						throw new IllegalStateException(
								"Unexpected cell cached formula result type: " + cell.getCachedFormulaResultType());

				}
				break;
			default:
				throw new RuntimeException("Unexpected cell type (" + cell.getCellType() + ")");

		}
		
        return text.toString();
	}

	private static void readHeaderFooter(HeaderFooter hf) {
		if (hf.getLeft() != null) {
			System.out.println(hf.getLeft());
		}
		if (hf.getCenter() != null) {
			System.out.println("\t");
			System.out.println(hf.getCenter());
		}
		if (hf.getRight() != null) {
			System.out.println("\t");
			System.out.println(hf.getRight());
		}

		System.out.println("\n");
	}

	private static String readSheet(Workbook wb, int iSheet) {
		Sheet sheet = wb.getSheetAt(iSheet);
		
		StringBuilder text = new StringBuilder();

		if (sheet != null) {
			System.out.println("Sheet " + sheet.getSheetName());

			readHeaderFooter(sheet.getHeader());

			int firstRow = sheet.getFirstRowNum();
			int lastRow = sheet.getLastRowNum(); // sheet.getPhysicalNumberOfRows();
			for (int j = firstRow; j <= lastRow; j++) {
				Row row = sheet.getRow(j);
				if (row != null) {
					int lastCell = row.getLastCellNum();

					for (int k = 0 /* row.getFirstCellNum() */ ; k < lastCell; k++) {
						Cell cell = row.getCell(k);
						if (cell != null) {
							System.out.print(getCellValue(cell));
							System.out.print(" | ");
						}
					}
					System.out.print("\n");
				}
			}

			readHeaderFooter(sheet.getFooter());

			System.out.println("\n\n" + "-".repeat(20));
		}

		return text.toString();

	}

	private static void readWorkbook(File file) throws IOException, InvalidFormatException {
		Workbook wb;
		byte[] bytes = readFileToBytes(file);
		FileMagic fileMagic = FileMagic.valueOf(bytes);
		switch (fileMagic) {
			case OOXML:
				wb = new XSSFWorkbook(file);
				break;
			case OLE2:
				wb = new HSSFWorkbook(new POIFSFileSystem(file));
				break;
			case BIFF4: // OldExcelExtractor
			default:
				throw new IllegalArgumentException("Received file does not have a standard excel extension.");

		}
		
		wb.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);

		
		System.out.println(getCellValue(wb.getSheetAt(0).getRow(1).getCell(2)));
		
		for (int iSheet = 0; iSheet < wb.getNumberOfSheets(); iSheet++) {
			readSheet(wb, iSheet);
		}
		
		wb.close();

	}

	private static byte[] readFileToBytes(File file) throws IOException {
		byte[] bytes = new byte[(int) file.length()];

		FileInputStream fis = null;
		try {

			fis = new FileInputStream(file);

			// read file into bytes[]
			fis.read(bytes);

		} finally {
			if (fis != null) {
				fis.close();
			}
		}

		return bytes;

	}
}
