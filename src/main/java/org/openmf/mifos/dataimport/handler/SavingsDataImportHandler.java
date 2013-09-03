package org.openmf.mifos.dataimport.handler;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Savings;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class SavingsDataImportHandler extends AbstractDataImportHandler {
	
	private static final Logger logger = LoggerFactory.getLogger(SavingsDataImportHandler.class);
	
	private static final int CLIENT_NAME_COL = 1;
    private static final int PRODUCT_COL = 2;
    private static final int FIELD_OFFICER_NAME_COL = 3;
    private static final int SUBMITTED_ON_DATE_COL = 4;
    private static final int APPROVED_DATE_COL = 5;
    private static final int ACTIVATION_DATE_COL = 6;
    private static final int NOMINAL_ANNUAL_INTEREST_RATE_COL = 7;
	private static final int INTEREST_COMPOUNDING_PERIOD_COL = 8;
	private static final int INTEREST_POSTING_PERIOD_COL = 9;
	private static final int INTEREST_CALCULATION_COL = 10;
	private static final int INTEREST_CALCULATION_DAYS_IN_YEAR_COL = 11;
	private static final int MIN_OPENING_BALANCE_COL = 12;
	private static final int LOCKIN_PERIOD_COL = 13;
	private static final int LOCKIN_PERIOD_FREQUENCY_COL = 14;
	private static final int WITHDRAWAL_FEE_AMOUNT_COL = 15;
	private static final int WITHDRAWAL_FEE_TYPE_COL = 16;
	private static final int ANNUAL_FEE_COL = 17;
	private static final int ANNUAL_FEE_ON_MONTH_DAY_COL = 18;
	private static final int STATUS_COL = 19;
    private static final int FAILURE_REPORT_COL = 20;

    private final RestClient restClient;
    
    private final Workbook workbook;
    
    private List<Savings> savings = new ArrayList<Savings>();

    public SavingsDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
    }
    
    @Override
    public Result parse() {
        Result result = new Result();
        Sheet savingsSheet = workbook.getSheet("Savings");
        Integer noOfEntries = getNumberOfRows(savingsSheet, 0);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = savingsSheet.getRow(rowIndex);
                String status = readAsString(STATUS_COL, row);
                if(status.equals("Imported"))
                	continue;
                String clientName = readAsString(CLIENT_NAME_COL, row);
                String clientId = getIdByName(workbook.getSheet("Clients"), clientName).toString();
                String productName = readAsString(PRODUCT_COL, row);
                String productId = getIdByName(workbook.getSheet("Products"), productName).toString();
                String fieldOfficerName = readAsString(FIELD_OFFICER_NAME_COL, row);
                String fieldOfficerId = getIdByName(workbook.getSheet("Staff"), fieldOfficerName).toString();
                String submittedOnDate = readAsDate(SUBMITTED_ON_DATE_COL, row);
                String approvalDate = readAsDate(APPROVED_DATE_COL, row);
                String activationDate = readAsDate(ACTIVATION_DATE_COL, row);
                String nominalAnnualInterestRate = readAsString(NOMINAL_ANNUAL_INTEREST_RATE_COL, row);
                String interestCompoundingPeriodType = readAsString(INTEREST_COMPOUNDING_PERIOD_COL, row);
                String interestCompoundingPeriodTypeId = "";
                if(interestCompoundingPeriodType.equals("Daily"))
                	interestCompoundingPeriodTypeId = "1";
                else if(interestCompoundingPeriodType.equals("Monthly"))
                	interestCompoundingPeriodTypeId = "4";
                String interestPostingPeriodType = readAsString(INTEREST_POSTING_PERIOD_COL, row);
                String interestPostingPeriodTypeId = "";
                if(interestPostingPeriodType.equals("Monthly"))
                	interestPostingPeriodTypeId = "4";
                else if(interestPostingPeriodType.equals("Quarterly"))
                	interestPostingPeriodTypeId = "5";
                else if(interestPostingPeriodType.equals("Annually"))
                	interestPostingPeriodTypeId = "7";
                String interestCalculationType = readAsString(INTEREST_CALCULATION_COL, row);
                String interestCalculationTypeId = "";
                if(interestCalculationType.equals("Daily Balance"))
                	interestCalculationTypeId = "1";
                else if(interestCalculationType.equals("Average Daily Balance"))
                	interestCalculationTypeId = "2";
                String interestCalculationDaysInYearType = readAsString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, row);
                String interestCalculationDaysInYearTypeId = "";
                if(interestCalculationDaysInYearType.equals("360 Days"))
                	interestCalculationDaysInYearTypeId = "360";
                else if(interestCalculationDaysInYearType.equals("365 Days"))
                	interestCalculationDaysInYearTypeId = "365";
                String minRequiredOpeningBalance = readAsString(MIN_OPENING_BALANCE_COL, row);
                String lockinPeriodFrequency = readAsString(LOCKIN_PERIOD_COL, row);
                String lockinPeriodFrequencyType = readAsString(LOCKIN_PERIOD_FREQUENCY_COL, row);
                String lockinPeriodFrequencyTypeId = "";
                if(lockinPeriodFrequencyType.equals("Days"))
                	lockinPeriodFrequencyTypeId = "0";
                else if(lockinPeriodFrequencyType.equals("Weeks"))
                	lockinPeriodFrequencyTypeId = "1";
                else if(lockinPeriodFrequencyType.equals("Months"))
                	lockinPeriodFrequencyTypeId = "2";
                else if(lockinPeriodFrequencyType.equals("Years"))
                	lockinPeriodFrequencyTypeId = "3";
                String withdrawalFeeAmount = readAsString(WITHDRAWAL_FEE_AMOUNT_COL, row);
                String withdrawalFeeType = readAsString(WITHDRAWAL_FEE_TYPE_COL, row);
                String withdrawalFeeTypeId = "";
                if(withdrawalFeeType.equals("Flat"))
                	withdrawalFeeTypeId = "1";
                else if(withdrawalFeeType.equals("% of Amount"))
                	withdrawalFeeTypeId = "2";
                String annualFeeAmount = readAsString(ANNUAL_FEE_COL, row);
                String annualFeeOnMonthDay = readAsDateWithoutYear(ANNUAL_FEE_ON_MONTH_DAY_COL, row);
                savings.add(new Savings(clientId, productId, fieldOfficerId, submittedOnDate, nominalAnnualInterestRate, interestCompoundingPeriodTypeId, interestPostingPeriodTypeId,
                		interestCalculationTypeId, interestCalculationDaysInYearTypeId, minRequiredOpeningBalance, lockinPeriodFrequency, lockinPeriodFrequencyTypeId, withdrawalFeeAmount,
                		withdrawalFeeTypeId, annualFeeAmount, annualFeeOnMonthDay, rowIndex, status));
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
        Sheet savingsSheet = workbook.getSheet("Savings");
        restClient.createAuthToken();
        for (Savings saving : savings) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(saving);
                logger.info(payload);
                restClient.post("savingsaccounts", payload);
                Cell statusCell = savingsSheet.getRow(saving.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (Exception e) {
            	String message = parseStatus(e.getMessage());
            	Row row = savingsSheet.getRow(saving.getRowIndex());
            	Cell statusCell = row.createCell(STATUS_COL);
            	statusCell.setCellValue(message);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
                Cell errorReportCell = row.createCell(FAILURE_REPORT_COL);
            	errorReportCell.setCellValue(message);
                result.addError("Row = " + saving.getRowIndex() + " ," + message);
            }
        }
        savingsSheet.setColumnWidth(STATUS_COL, 4000);
    	writeString(STATUS_COL, savingsSheet.getRow(0), "Status");
        return result;
    }
    
}
