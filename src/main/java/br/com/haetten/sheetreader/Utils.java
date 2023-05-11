package br.com.haetten.sheetreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Utils {

	public static byte[] readFileToBytes(File file) throws IOException {
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
