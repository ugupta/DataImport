package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.openmf.mifos.dataimport.dto.GeneralClient;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class LoanWorkbookPopulator extends AbstractWorkbookPopulator {
	
	private static final Logger logger = LoggerFactory.getLogger(LoanWorkbookPopulator.class);
	
	private final RestClient restClient;
	
	private ClientSheetPopulator clientSheetPopulator;
	private PersonnelSheetPopulator personnelSheetPopulator;
	private ProductSheetPopulator productSheetPopulator;
	private FundSheetPopulator fundSheetPopulator;
	
	public static final int OFFICE_NAME_COL = 0;
    public static final int CLIENT_NAME_COL = 1;
    public static final int PRODUCT_COL = 2;
    public static final int LOAN_OFFICER_NAME_COL = 3;
    public static final int SUBMITTED_ON_DATE_COL = 4;
    public static final int FUND_NAME_COL = 5;
    public static final int LOOKUP_CLIENT_NAME_COL = 26;
    public static final int LOOKUP_ACTIVATION_DATE_COL = 27;
	
	public LoanWorkbookPopulator(RestClient restClient) {
    	this.restClient = restClient;
    }
	
	 @Override
	    public Result downloadAndParse() {
	    	clientSheetPopulator = new ClientSheetPopulator(restClient);
	    	personnelSheetPopulator = new PersonnelSheetPopulator(Boolean.TRUE, restClient);
	    	productSheetPopulator = new ProductSheetPopulator(restClient);
	    	fundSheetPopulator = new FundSheetPopulator(restClient);
	    	Result result =  clientSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = personnelSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = productSheetPopulator.downloadAndParse();
	    	if(result.isSuccess())
	    		result = fundSheetPopulator.downloadAndParse();
	    	return result;
	    }

	    @Override
	    public Result populate(Workbook workbook) {
	    	Sheet loanSheet = workbook.createSheet("Loans");
	    	Result result = clientSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = personnelSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = productSheetPopulator.populate(workbook);
	    	if(result.isSuccess())
	    		result = fundSheetPopulator.populate(workbook);
	    	setLayout(loanSheet);
	    	setDateLookupTable(workbook, loanSheet);
	    	if(result.isSuccess())
	            result = setRules(loanSheet);
	        return result;
	    }
	    
	    private void setLayout(Sheet worksheet) {
	    	Row rowHeader = worksheet.createRow(0);
	        rowHeader.setHeight((short)500);
	        worksheet.setColumnWidth(OFFICE_NAME_COL, 6000);
            worksheet.setColumnWidth(CLIENT_NAME_COL, 6000);
            worksheet.setColumnWidth(PRODUCT_COL, 4000);
            worksheet.setColumnWidth(LOAN_OFFICER_NAME_COL, 6000);
            worksheet.setColumnWidth(SUBMITTED_ON_DATE_COL, 4000);
            worksheet.setColumnWidth(FUND_NAME_COL, 4000);
            worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 6000);
            worksheet.setColumnWidth(LOOKUP_ACTIVATION_DATE_COL, 6000);
            writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
            writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
            writeString(PRODUCT_COL, rowHeader, "Product*");
            writeString(LOAN_OFFICER_NAME_COL, rowHeader, "Loan Officer*");
            writeString(SUBMITTED_ON_DATE_COL, rowHeader, "Submitted On*");
            writeString(FUND_NAME_COL, rowHeader, "Fund Name");
            writeString(LOOKUP_CLIENT_NAME_COL, rowHeader, "Client Name");
            writeString(LOOKUP_ACTIVATION_DATE_COL, rowHeader, "Client Activation Date");
	    }
	    
	    private Result setDateLookupTable(Workbook workbook, Sheet loanSheet) {
	    	Result result = new Result();
	    	try {
	    	CellStyle dateCellStyle = workbook.createCellStyle();
	        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
	        dateCellStyle.setDataFormat(df);	
	    	int rowIndex = 0;	
	    	List<GeneralClient> clients = clientSheetPopulator.getClients();
    		for(GeneralClient client: clients) {
    			Row row = loanSheet.createRow(++rowIndex);
    			writeString(LOOKUP_CLIENT_NAME_COL, row, client.getDisplayName().replaceAll("[ )(] ", "_"));
    			writeDate(LOOKUP_ACTIVATION_DATE_COL, row, client.getActivationDate().get(2) + "/" + client.getActivationDate().get(1) + "/" + client.getActivationDate().get(0), dateCellStyle);
    		}
	    	} catch (Exception e) {
	    		result.addError(e.getMessage());
	    		logger.error(e.getMessage());
	    	}
	    	return result;
	    }
	    
	    private Result setRules(Sheet worksheet) {
	    	Result result = new Result();
	    	try {
	    		CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), OFFICE_NAME_COL, OFFICE_NAME_COL);
	        	CellRangeAddressList clientNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), CLIENT_NAME_COL, CLIENT_NAME_COL);
	        	CellRangeAddressList productNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PRODUCT_COL, PRODUCT_COL);
	        	CellRangeAddressList loanOfficerRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_OFFICER_NAME_COL, LOAN_OFFICER_NAME_COL);
	        	CellRangeAddressList dateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), SUBMITTED_ON_DATE_COL, SUBMITTED_ON_DATE_COL);
	        	CellRangeAddressList fundNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), FUND_NAME_COL, FUND_NAME_COL);
	        	
	        	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
	        	Workbook loanWorkbook = worksheet.getWorkbook();
	        	
	        	Name officeGroup = loanWorkbook.createName();
	        	officeGroup.setNameName("Office");
	        	officeGroup.setRefersToFormula("Clients!$A$2:$A$" + (clientSheetPopulator.officeNames.size() + 1));
	        	
	        	ArrayList<String> officeNames = clientSheetPopulator.officeNames;
	        	Name[] clientGroups = new Name[officeNames.size()];
	        	ArrayList<String> formulas = new ArrayList<String>();
	        	for(Integer i = 0, j = 2; i < officeNames.size(); i++, j = j + 2) {
	        		String lastColumnLetters = CellReference.convertNumToColString(clientSheetPopulator.lastColumnLetters.get(i));
	        		formulas.add("Clients!$B$" + j + ":$" + lastColumnLetters + "$" + j);
	        		clientGroups[i] = loanWorkbook.createName();
	        	    clientGroups[i].setNameName(officeNames.get(i));
	        	    clientGroups[i].setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	Name productGroup = loanWorkbook.createName();
	        	productGroup.setNameName("Products");
	        	productGroup.setRefersToFormula("Products!$B$2:$B$" + (productSheetPopulator.getProductsSize() + 1));
	        	
	        	Name[] loanOfficerGroups = new Name[officeNames.size()];
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0, j = 2; i < officeNames.size(); i++, j = j + 2) {
	        		String lastColumnLetters = CellReference.convertNumToColString(personnelSheetPopulator.lastColumnLetters.get(i));
	        		formulas.add("Staff!$B$" + j + ":$" + lastColumnLetters + "$" + j);
	        		loanOfficerGroups[i] = loanWorkbook.createName();
	        	    loanOfficerGroups[i].setNameName(officeNames.get(i)+"X");
	        	    loanOfficerGroups[i].setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	Name fundGroup = loanWorkbook.createName();
	        	fundGroup.setNameName("Funds");
	        	fundGroup.setRefersToFormula("Funds!$B$2:$B$" + (fundSheetPopulator.getFundsSize() + 1));
	        	
	        	DataValidationConstraint officeNameConstraint = validationHelper.createExplicitListConstraint(clientSheetPopulator.getOfficeNames());
	        	DataValidationConstraint clientNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT($A1)");
	        	DataValidationConstraint productNameConstraint = validationHelper.createFormulaListConstraint("Products");
	        	DataValidationConstraint loanOfficerNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE($A1,\"X\"))");
	        	DataValidationConstraint activationDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=VLOOKUP($B1,$AA$2:$AB$" + (clientSheetPopulator.getClientsSize() + 1) + ",2,FALSE)", "=TODAY()", "dd/mm/yy");
	        	DataValidationConstraint fundNameConstraint = validationHelper.createFormulaListConstraint("Funds");
	        	
	        	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
	        	officeValidation.setSuppressDropDownArrow(false);
	        	DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
	        	clientValidation.setSuppressDropDownArrow(false);
	        	DataValidation productNameValidation = validationHelper.createValidation(productNameConstraint, productNameRange);
	        	productNameValidation.setSuppressDropDownArrow(false);
	        	DataValidation loanOfficerValidation = validationHelper.createValidation(loanOfficerNameConstraint, loanOfficerRange);
	        	loanOfficerValidation.setSuppressDropDownArrow(false);
	        	DataValidation activationDateValidation = validationHelper.createValidation(activationDateConstraint, dateRange);
	        	activationDateValidation.setSuppressDropDownArrow(false);
	        	DataValidation fundNameValidation = validationHelper.createValidation(fundNameConstraint, fundNameRange);
	        	fundNameValidation.setSuppressDropDownArrow(false);
	        	
	        	worksheet.addValidationData(officeValidation);
	            worksheet.addValidationData(clientValidation);
	            worksheet.addValidationData(productNameValidation);
	            worksheet.addValidationData(loanOfficerValidation);
	            worksheet.addValidationData(activationDateValidation);
	            worksheet.addValidationData(fundNameValidation);
	    	} catch (Exception e) {
	    		result.addError(e.getMessage());
	    	}
	       return result;
	    }
}
