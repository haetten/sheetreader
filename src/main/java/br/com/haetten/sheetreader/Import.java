package br.com.haetten.sheetreader;

import java.util.Date;

import lombok.Data;

@Data
public class Import {
	Integer order;
	String brPatentApplication;
	String originalApplication;
	String patBaseLink;
	Date filingDate;
	Date publicationDate;
	String applicant;
	String title;
	String abstract_;
	String publishedApplication;
	byte[] publishedApplicationFile;
}
