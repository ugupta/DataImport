package org.openmf.mifos.dataimport.populator;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public abstract class AbstractWorkbookPopulator implements WorkbookPopulator {

	    protected void writeInt(int colIndex, Row row, int value) {
	            row.createCell(colIndex).setCellValue(value);
	    }

	    protected void writeString(int colIndex, Row row, String value) {
	            row.createCell(colIndex).setCellValue(value);
	    }
	    
	    protected void writeDouble(int colIndex, Row row, double value) {
	    	    row.createCell(colIndex).setCellValue(value);
	    }

	    protected void writeFormula(int colIndex, Row row, String formula) {
	    	    row.createCell(colIndex).setCellType(Cell.CELL_TYPE_FORMULA);
	    	    row.createCell(colIndex).setCellFormula(formula);
	    }
	    protected void writeDate(int colIndex, Row row, String value, CellStyle dateCellStyle) {
	    	try {
	    		//To make validation between functions inclusive.
	    	    Date date = new SimpleDateFormat("dd/MM/yyyy", Locale.ENGLISH).parse(value);
	    		Calendar cal = Calendar.getInstance();
	    		cal.setTime(date);
	    	    cal.set(Calendar.HOUR_OF_DAY, 0);
	    	    cal.set(Calendar.MINUTE, 0);
	    	    cal.set(Calendar.SECOND, 0);
	    	    cal.set(Calendar.MILLISECOND, 0);
	    	    Date dateWithoutTime = cal.getTime();
	    	    row.createCell(colIndex).setCellValue(dateWithoutTime);
	    	    row.getCell(colIndex).setCellStyle(dateCellStyle);
	    	    } catch (Exception e) {
	    	    	throw new IllegalArgumentException("ParseException");
	    	    }
	    }
}
