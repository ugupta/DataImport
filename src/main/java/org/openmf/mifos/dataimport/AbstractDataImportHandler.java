package org.openmf.mifos.dataimport;

import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public abstract class AbstractDataImportHandler implements DataImportHandler {

    protected final Sheet sheet;

    public AbstractDataImportHandler(Sheet sheet) {
        this.sheet = sheet;
    };

    protected Integer getNumberOfRows() {
        Integer noOfEntries = 1;
        // getLastRowNum and getPhysicalNumberOfRows showing false values
        // sometimes.
        while (sheet.getRow(noOfEntries) != null) {
            noOfEntries++;
        }
        return noOfEntries;
    }

    protected int readAsInt(int colIndex, Row row) {
        return ((Double) row.getCell(colIndex).getNumericCellValue()).intValue();
    }

    protected String readAsString(int colIndex, Row row) {
        try {
            return row.getCell(colIndex).getStringCellValue();
        } catch (Exception e) {
            return row.getCell(colIndex).getNumericCellValue() + "";
        }
    }

    protected Date readAsDate(int colIndex, Row row) {
        return row.getCell(colIndex).getDateCellValue();
    }

}
