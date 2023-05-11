package br.com.haetten.sheetreader;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.text.ParseException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TechMappingSheetReader {

	private static Workbook readWorkbook(File file) throws InvalidFormatException, IOException {
		Workbook wb;
		byte[] bytes = Utils.readFileToBytes(file);
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
		
		return wb;
		
	}
	
	public ImportExcelFile leFormatoEspecifico(File file) throws InvalidFormatException, IOException, ParseException {
		Workbook wb = readWorkbook(file);
		Sheet sheet = wb.getSheetAt(0);
		
		ImportExcelFile importa = new ImportExcelFile();
		importa.setTitle(sheet.getRow(0).getCell(0).toString());
		importa.setFile(Utils.readFileToBytes(file));
		importa.setImportDate(new Date());
		
		int firstRow = 2;
		int lastRow = sheet.getLastRowNum(); // sheet.getPhysicalNumberOfRows();
		for (int j = firstRow; j <= lastRow; j++) {
			Row row = sheet.getRow(j);
			if (row != null) {
				if(row.getCell(0) == null || row.getCell(0).toString() == null) {
					continue;
				}
				
				Import i = new Import();
				Cell cell;
				cell = row.getCell(0); i.setOrder(cell == null ? null : (int) cell.getNumericCellValue());

				cell = row.getCell(1); i.setBrPatentApplication(cell == null ? null : cell.toString());
				
				cell = row.getCell(2);
				i.setOriginalApplication(cell == null ? null : cell.toString());
				i.setPatBaseLink(cell == null || cell.getHyperlink() == null ? null : cell.getHyperlink().getAddress());
				
				cell = row.getCell(3); i.setFilingDate(cell == null ? null : cell.getDateCellValue());
				cell = row.getCell(4); i.setPublicationDate(cell == null ? null : cell.getDateCellValue());
				
				cell = row.getCell(5); i.setApplicant(cell == null ? null : cell.toString());
				cell = row.getCell(6); i.setTitle(cell == null ? null : cell.toString());
				cell = row.getCell(7); i.setAbstract_(cell == null ? null : cell.toString());
				
				cell = row.getCell(8);
				i.setPublishedApplication(cell == null || cell.getHyperlink() == null ? null : cell.getHyperlink().getAddress());
				if (i.getPublishedApplication() != null) {
					URL url = new URL(i.getPublishedApplication());
					ByteArrayOutputStream baos = new ByteArrayOutputStream();
					InputStream is = null;
					try {
						is = url.openStream();
						byte[] byteChunk = new byte[4096];
						int n;

						while ((n = is.read(byteChunk)) > 0) {
							baos.write(byteChunk, 0, n);
						}
						
						i.setPublishedApplicationFile(byteChunk);
						
					} catch (IOException e) {
						System.err.printf("Failed while reading bytes from %s: %s", url.toExternalForm(), e.getMessage());
						e.printStackTrace();
					} finally {
						if (is != null) {
							is.close();
						}
					}
				}
				
				
				importa.getImports().add(i);

			}
		}
		
		System.out.println(importa);
		
		return importa;

	}
}
