package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.CompactLoan;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

public class LoanSheetPopulator extends AbstractWorkbookPopulator {
	
private static final Logger logger = LoggerFactory.getLogger(LoanSheetPopulator.class);
	
    private final RestClient restClient;

    private String content;
    
    private List<CompactLoan> loans = new ArrayList<CompactLoan>();
    
    private static final int CLIENT_NAME_COL = 0;
    private static final int ACCOUNT_NO_COL = 1;
    private static final int PRODUCT_COL = 2;
    private static final int PRINCIPAL_COL = 3;
    
    public LoanSheetPopulator(RestClient restClient) {
    	this.restClient = restClient;
    }
    
    @Override
    public Result downloadAndParse() {
    	Result result = new Result();
    	try {
        	restClient.createAuthToken();
            content = restClient.get("loans");
            Gson gson = new Gson();
            JsonParser parser = new JsonParser();
            JsonObject obj = parser.parse(content).getAsJsonObject();
            JsonArray array = obj.getAsJsonArray("pageItems");
            Iterator<JsonElement> iterator = array.iterator();
            while(iterator.hasNext()) {
            	JsonElement json = iterator.next();
            	CompactLoan loan = gson.fromJson(json, CompactLoan.class);
            	if(loan.isActive())
            	  loans.add(loan);
            }
            logger.info(loans.size()+" ");
       } catch (Exception e) {
           result.addError(e.getMessage());
           logger.error(e.getMessage());
       }
	return result;
    }
    
    @Override
    public Result populate(Workbook workbook) {
    	Result result = new Result();
    	Sheet loanSheet = workbook.createSheet("Loans");
    	setLayout(loanSheet);
    	int rowIndex = 1;
    	Row row;
    	Collections.sort(loans, CompactLoan.ClientNameComparator);
    	try{
    		for(CompactLoan loan : loans) {
    			row = loanSheet.createRow(rowIndex++);
    			writeString(CLIENT_NAME_COL, row, loan.getClientName().replaceAll("[ )(]", "_"));
    			writeInt(ACCOUNT_NO_COL, row, Integer.parseInt(loan.getAccountNo()));
    			writeString(PRODUCT_COL, row, loan.getLoanProductName());
    			writeDouble(PRINCIPAL_COL, row, loan.getPrincipal());
    		}
	   } catch (Exception e) {
		result.addError(e.getMessage());
		logger.error(e.getMessage());
	    }
	
        return result;
    }

	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        worksheet.setColumnWidth(CLIENT_NAME_COL, 5000);
        worksheet.setColumnWidth(ACCOUNT_NO_COL, 3000);
        worksheet.setColumnWidth(PRODUCT_COL, 4000);
        worksheet.setColumnWidth(PRINCIPAL_COL, 4000);
        writeString(CLIENT_NAME_COL, rowHeader, "Client Name");
        writeString(ACCOUNT_NO_COL, rowHeader, "Account No.");
        writeString(PRODUCT_COL, rowHeader, "Product Name");
        writeString(PRINCIPAL_COL, rowHeader, "Principal");
	}

}
