package org.openmf.mifos.dataimport.populator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class LoanWorkbookPopulator extends AbstractWorkbookPopulator {
	
	private static final Logger logger = LoggerFactory.getLogger(LoanWorkbookPopulator.class);
	
	private final RestClient client;
	
	public LoanWorkbookPopulator(RestClient client) {
    	this.client = client;
    }
	
	 @Override
	    public Result downloadAndParse() {
	    	Result result = new Result();
	    	
	    	return result;
	    }

	    @Override
	    public Result populate(Workbook workbook) {
	    	Sheet clientSheet = workbook.createSheet("Loans");
	    	Result result = new Result();
	        return result;
	    }

}
