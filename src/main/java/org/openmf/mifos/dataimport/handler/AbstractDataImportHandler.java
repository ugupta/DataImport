package org.openmf.mifos.dataimport.handler;

import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public abstract class AbstractDataImportHandler implements DataImportHandler {

    protected Integer getNumberOfRows(Sheet sheet) {
        Integer noOfEntries = 1;
        // getLastRowNum and getPhysicalNumberOfRows showing false values
        // sometimes.
        while (sheet.getRow(noOfEntries).getCell(0) != null) {
            noOfEntries++;
        }
        return noOfEntries;
    }

    protected int readAsInt(int colIndex, Row row) {
        return ((Double) row.getCell(colIndex).getNumericCellValue()).intValue();
    }

    protected String readAsString(int colIndex, Row row) {
        try {
        	Cell c = row.getCell(colIndex);
        	if (c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
        		return "";
            return c.getStringCellValue();
        } catch (Exception e) {
            return ((Double)row.getCell(colIndex).getNumericCellValue()).intValue() + "";
        }
    }

    protected String readAsDate(int colIndex, Row row) {
    	try{
    		Cell c = row.getCell(colIndex);
    		if(c == null || c.getCellType() == Cell.CELL_TYPE_BLANK)
    			return "";
    		DateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");
            return dateFormat.format(c.getDateCellValue());
    	}  catch  (Exception e) {
    		return e.getMessage();
    	}
    	
    }
    
    protected Boolean readAsBoolean(int colIndex, Row row) {
    	return row.getCell(colIndex).getBooleanCellValue();
    }
    
    public Integer getIdByName (Sheet sheet, String name) {
    	String sheetName = sheet.getSheetName();
    	if(sheetName.equals("Offices") || sheetName.equals("Clients") || sheetName.equals("Staff") || sheetName.equals("Extras")) {
    	for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(name)) {
                    	if(sheet.getSheetName().equals("Offices"))
                            return ((Double)row.getCell(cell.getColumnIndex() - 1).getNumericCellValue()).intValue(); 
                    	else if(sheet.getSheetName().equals("Staff") || sheet.getSheetName().equals("Clients"))
                           return ((Double)sheet.getRow(row.getRowNum() + 1).getCell(cell.getColumnIndex()).getNumericCellValue()).intValue();
                    	else if(sheet.getSheetName().equals("Extras"))
                    		return ((Double)row.getCell(cell.getColumnIndex() - 1).getNumericCellValue()).intValue();
                    }
                }
            }
          }
    	} else if (sheetName.equals("Products")) {
    		for(Row row : sheet) {
    			for(int i = 0; i < 2; i++) {
    				Cell cell = row.getCell(i);
    				if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
    					if(cell.getRichStringCellValue().getString().trim().equals(name)) {
    						return ((Double)row.getCell(cell.getColumnIndex() - 1).getNumericCellValue()).intValue();
    					}
    				}
    			}
    		}
    	}
        return 0;
    }

}
