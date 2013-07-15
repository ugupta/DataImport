package org.openmf.mifos.dataimport.populator;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public abstract class AbstractWorkbookPopulator implements WorkbookPopulator {

	    protected void writeInt(int colIndex, Row row, int value) {
	            row.createCell(colIndex).setCellValue(value);
	    }

	    protected void writeString(int colIndex, Row row, String value) {
	            row.createCell(colIndex).setCellValue(value);
	    }

	    protected void writeDate(int colIndex, Row row, String value, CellStyle dateCellStyle) {
	    	try {
	    	    Date date = new SimpleDateFormat("mm/dd/yyyy", Locale.ENGLISH).parse(value);
	    	    row.createCell(colIndex).setCellValue(date);
	    	    row.getCell(colIndex).setCellStyle(dateCellStyle);
	    	    } catch (Exception e) {
	    	    	throw new IllegalArgumentException("ParseException");
	    	    }
	    }
}
