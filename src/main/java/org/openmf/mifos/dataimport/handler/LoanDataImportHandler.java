package org.openmf.mifos.dataimport.handler;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Loan;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class LoanDataImportHandler extends AbstractDataImportHandler {
	
	private static final Logger logger = LoggerFactory.getLogger(LoanDataImportHandler.class);
	
	public static final int OFFICE_NAME_COL = 0;
    public static final int CLIENT_NAME_COL = 1;
    public static final int PRODUCT_COL = 2;
    public static final int LOAN_OFFICER_NAME_COL = 3;
    public static final int SUBMITTED_ON_DATE_COL = 4;
    public static final int FUND_NAME_COL = 5;
    public static final int PRINCIPAL_COL = 6;
    public static final int NO_OF_REPAYMENTS_COL = 7;
    public static final int REPAID_EVERY_COL = 8;
    public static final int REPAID_EVERY_FREQUENCY_COL = 9;
    public static final int LOAN_TERM_COL = 10;
    public static final int LOAN_TERM_FREQUENCY_COL = 11;
    public static final int NOMINAL_INTEREST_RATE_COL = 12;
    public static final int NOMINAL_INTEREST_RATE_FREQUENCY_COL = 13;
    public static final int DISBURSED_DATE_COL = 14;
    public static final int AMORTIZATION_COL = 15;
    public static final int INTEREST_METHOD_COL = 16;
    public static final int INTEREST_CALCULATION_PERIOD_COL = 17;
    public static final int ARREARS_TOLERANCE_COL = 18;
    public static final int REPAYMENT_STRATEGY_COL = 19;
    public static final int GRACE_ON_PRINCIPAL_PAYMENT_COL = 20;
    public static final int GRACE_ON_INTEREST_PAYMENT_COL = 21;
    public static final int GRACE_ON_INTEREST_CHARGED_COL = 22;
    public static final int INTEREST_CHARGED_FROM_COL = 23;
    public static final int FIRST_REPAYMENT_COL = 24;
    
    private List<Loan> loans = new ArrayList<Loan>();
    
    private final RestClient restClient;
    
    private final Workbook workbook;

    public LoanDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
    }
    
    @Override
    public Result parse() {
        Result result = new Result();
        Sheet loanSheet = workbook.getSheet("Loans");
        Integer noOfEntries = getNumberOfRows(loanSheet);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = loanSheet.getRow(rowIndex);
                String clientName = readAsString(CLIENT_NAME_COL, row);
                String productName = readAsString(PRODUCT_COL, row);
                String loanOfficerName = readAsString(LOAN_OFFICER_NAME_COL, row);
                String submittedOnDate = readAsDate(SUBMITTED_ON_DATE_COL, row);
                String disbursedDate = readAsDate(DISBURSED_DATE_COL, row);
                String interestChargedFromDate = readAsDate(INTEREST_CHARGED_FROM_COL, row);
                String firstRepaymentOnDate = readAsDate(FIRST_REPAYMENT_COL, row);
            } catch (Exception e) {
                logger.error("row = " + rowIndex, e);
                result.addError("Row = " + rowIndex + " , " + e.getMessage());
            }
        }
    
     return result;
   }
    
    @Override
    public Result upload() {
        Result result = new Result();
        restClient.createAuthToken();
        for (Loan loan : loans) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(loan);
                logger.info(payload);
                restClient.post("loans", payload);
            } catch (Exception e) {
            }
        }
        return result;
    }
}
