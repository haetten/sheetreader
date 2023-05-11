package br.com.haetten.sheetreader;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import lombok.Data;

@Data
public class ImportExcelFile {
	List<Import> imports = new ArrayList<>();
	byte[] file;
	Date importDate;
	String title;
}
