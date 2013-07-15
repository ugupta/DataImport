package org.openmf.mifos.dataimport;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Row;

public abstract class AbstractWorkbookPopulator implements WorkbookPopulator {

	    protected Row writeInt(int colIndex, Row row, int value) {
	            row.getCell(colIndex).setCellValue(value);
	    	    return row;
	    }

	    protected Row writeString(int colIndex, Row row, String value) {
	            row.getCell(colIndex).setCellValue(value);
	            return row;
	    }

	    protected Row writeDate(int colIndex, Row row, String value)throws ParseException {
	    	    Date date = new SimpleDateFormat("dd MMMM, yyyy", Locale.ENGLISH).parse(value);
	    	    row.getCell(colIndex).setCellValue(date);
	    	    return row;
	    }
}
