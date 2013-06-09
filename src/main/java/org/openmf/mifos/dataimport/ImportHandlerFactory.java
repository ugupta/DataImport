package org.openmf.mifos.dataimport;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;


public class ImportHandlerFactory {
    
    
    public static final DataImportHandler createImportHandler(InputStream content) throws IOException {
        
        DataImportHandler importHandler = null;

        Workbook workbook = new HSSFWorkbook(content);
        
        if(workbook.getSheet("Offices") != null) {
            importHandler = new OfficeDataImportHandler(workbook.getSheet("Offices"));
        }
        
        return importHandler;
    }

}
