package org.openmf.mifos.dataimport;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openmf.mifos.dataimport.dto.Office;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class OfficeDataImportHandler extends AbstractDataImportHandler {

    public static final int EXTERNAL_ID_COL = 0;
    public static final int PARENT_ID_COL = 1;
    public static final int OFFICE_NAME_COL = 2;
    public static final int OPENING_DATE_COL = 3;

    private static final Logger logger = LoggerFactory.getLogger(OfficeDataImportHandler.class);

    private List<Office> offices = new ArrayList<Office>();

    public OfficeDataImportHandler(Sheet sheet) {
        super(sheet);
    }

    @Override
    public Result parse() {
        Result result = new Result();
        Integer noOfEntries = getNumberOfRows();
        logger.info(noOfEntries.toString());
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            try {
                Row row = sheet.getRow(rowIndex);
                String externalId = readAsString(EXTERNAL_ID_COL, row);
                int parentId = readAsInt(PARENT_ID_COL, row);
                String officeName = readAsString(OFFICE_NAME_COL, row);
                Date openingDate = readAsDate(OPENING_DATE_COL, row);
                logger.debug("Row Contents: " + parentId + ", " + externalId + ", " + officeName + ", " + openingDate);
                offices.add(new Office(officeName, parentId, externalId, openingDate, rowIndex));
            } catch (Exception e) {
                logger.error("row = " + rowIndex, e);
                result.addError("Row = " + rowIndex + " ," + e.getMessage());
            }
        }

        return null;
    }

    @Override
    public Result upload() {
        Result result = new Result();
        for (Office office : offices) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(office);
                post("offices", payload);
            } catch (Exception e) {
                logger.error("row = " + office.getRowIndex(), e);
                result.addError("Row = " + office.getRowIndex() + " ," + e.getMessage());
            }

        }
        return null;
    }

}
