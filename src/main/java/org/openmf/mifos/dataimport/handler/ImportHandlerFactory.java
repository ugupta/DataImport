package org.openmf.mifos.dataimport.handler;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.handler.client.ClientDataImportHandler;
import org.openmf.mifos.dataimport.handler.loan.LoanDataImportHandler;
import org.openmf.mifos.dataimport.handler.loan.LoanRepaymentDataImportHandler;
import org.openmf.mifos.dataimport.handler.savings.SavingsDataImportHandler;
import org.openmf.mifos.dataimport.handler.savings.SavingsTransactionDataImportHandler;
import org.openmf.mifos.dataimport.http.MifosRestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class ImportHandlerFactory {
	
	private static final Logger logger = LoggerFactory.getLogger(ImportHandlerFactory.class);
    
    public static final DataImportHandler createImportHandler(Workbook workbook) throws IOException {
        
        if(workbook.getSheetIndex("Clients") == 0) {
        	    logger.info("CROSSCHECK");
            	return new ClientDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("Loans") == 0) {
        	    return new LoanDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("LoanRepayment") == 0) {
        	    return new LoanRepaymentDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("Savings") == 0) {
    	    return new SavingsDataImportHandler(workbook, new MifosRestClient());
        } else if(workbook.getSheetIndex("SavingsTransaction") == 0) {
    	    return new SavingsTransactionDataImportHandler(workbook, new MifosRestClient());
        }
        throw new IllegalArgumentException("No work sheet found for processing : active sheet " + workbook.getSheetName(0));
    }

}
