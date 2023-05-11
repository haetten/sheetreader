package br.com.haetten.sheetreader;

import java.io.IOException;
import java.text.ParseException;

import javax.swing.JFileChooser;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Main {
	public static void main(String... args) throws InvalidFormatException, ParseException {
		JFileChooser fc = new JFileChooser();

		if (fc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
			try {
//				Workbook wb = readWorkbook(fc.getSelectedFile());
//				printWorkbook(wb);
				System.out.println(new TechMappingSheetReader().leFormatoEspecifico(fc.getSelectedFile()));
				
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}
	

	private static String getCellValue(Cell cell) {
        StringBuilder text = new StringBuilder();
        HSSFDataFormatter _formatter = new HSSFDataFormatter();
        
        if(cell == null) {
        	return "";
        }
        
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
	

	private static void printWorkbook(Workbook wb) throws IOException, InvalidFormatException {		
		for (int iSheet = 0; iSheet < wb.getNumberOfSheets(); iSheet++) {
			readSheet(wb, iSheet);
		}
		
		wb.close();

	}
}
