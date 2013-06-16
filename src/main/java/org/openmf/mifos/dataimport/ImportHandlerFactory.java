package org.openmf.mifos.dataimport;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.http.MifosRestClient;


public class ImportHandlerFactory {
    
    
    public static final DataImportHandler createImportHandler(InputStream content, @SuppressWarnings("unused") ImportFormatType type) throws IOException {

        Workbook workbook = new HSSFWorkbook(content);
        
        if(workbook.getSheet("Offices") != null) {
             return new OfficeDataImportHandler(workbook.getSheet("Offices"), new MifosRestClient());
        }
        throw new IllegalArgumentException("No work sheet found for processing : active sheet " + workbook.getSheetName(0));
    }

}
