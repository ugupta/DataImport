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
import org.openmf.mifos.dataimport.dto.Product;
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
	    	if(result.isSuccess())
	            result = setRules(loanSheet);
	    	if(result.isSuccess())
	            result = setDefaults(loanSheet);
	    	setDateLookupTable(workbook, loanSheet);
	        return result;
	    }
	    
	    private void setLayout(Sheet worksheet) {
	    	Row rowHeader = worksheet.createRow(0);
	        rowHeader.setHeight((short)500);
	        worksheet.setColumnWidth(OFFICE_NAME_COL, 4000);
            worksheet.setColumnWidth(CLIENT_NAME_COL, 4000);
            worksheet.setColumnWidth(PRODUCT_COL, 4000);
            worksheet.setColumnWidth(LOAN_OFFICER_NAME_COL, 4000);
            worksheet.setColumnWidth(SUBMITTED_ON_DATE_COL, 3200);
            worksheet.setColumnWidth(FUND_NAME_COL, 4000);
            worksheet.setColumnWidth(PRINCIPAL_COL, 3000);
            worksheet.setColumnWidth(LOAN_TERM_COL, 2000);
            worksheet.setColumnWidth(LOAN_TERM_FREQUENCY_COL, 2500);
            worksheet.setColumnWidth(NO_OF_REPAYMENTS_COL, 3800);
            worksheet.setColumnWidth(REPAID_EVERY_COL, 2000);
            worksheet.setColumnWidth(REPAID_EVERY_FREQUENCY_COL, 2000);
            worksheet.setColumnWidth(NOMINAL_INTEREST_RATE_COL, 2000);
            worksheet.setColumnWidth(NOMINAL_INTEREST_RATE_FREQUENCY_COL, 3000);
            worksheet.setColumnWidth(DISBURSED_DATE_COL, 3700);
            worksheet.setColumnWidth(AMORTIZATION_COL, 6000);
            worksheet.setColumnWidth(INTEREST_METHOD_COL, 4500);
            worksheet.setColumnWidth(INTEREST_CALCULATION_PERIOD_COL, 6000);
            worksheet.setColumnWidth(ARREARS_TOLERANCE_COL, 4000);
            worksheet.setColumnWidth(REPAYMENT_STRATEGY_COL, 4700);
            worksheet.setColumnWidth(GRACE_ON_PRINCIPAL_PAYMENT_COL, 5000);
            worksheet.setColumnWidth(GRACE_ON_INTEREST_PAYMENT_COL, 5000);
            worksheet.setColumnWidth(GRACE_ON_INTEREST_CHARGED_COL, 5000);
            worksheet.setColumnWidth(INTEREST_CHARGED_FROM_COL, 4000);
            worksheet.setColumnWidth(FIRST_REPAYMENT_COL, 4700);
            worksheet.setColumnWidth(LOOKUP_CLIENT_NAME_COL, 6000);
            worksheet.setColumnWidth(LOOKUP_ACTIVATION_DATE_COL, 6000);
            writeString(OFFICE_NAME_COL, rowHeader, "Office Name*");
            writeString(CLIENT_NAME_COL, rowHeader, "Client Name*");
            writeString(PRODUCT_COL, rowHeader, "Product*");
            writeString(LOAN_OFFICER_NAME_COL, rowHeader, "Loan Officer*");
            writeString(SUBMITTED_ON_DATE_COL, rowHeader, "Submitted On*");
            writeString(FUND_NAME_COL, rowHeader, "Fund Name");
            writeString(PRINCIPAL_COL, rowHeader, "Principal*");
            writeString(LOAN_TERM_COL, rowHeader, "Loan Term*");
            writeString(NO_OF_REPAYMENTS_COL, rowHeader, "# of Repayments*");
            writeString(REPAID_EVERY_COL, rowHeader, "Repaid Every*");
            writeString(NOMINAL_INTEREST_RATE_COL, rowHeader, "Nominal Interest %*");
            writeString(DISBURSED_DATE_COL, rowHeader, "Disbursed Date*");
            writeString(AMORTIZATION_COL, rowHeader, "Amortization*");
            writeString(INTEREST_METHOD_COL, rowHeader, "Interest Method*");
            writeString(INTEREST_CALCULATION_PERIOD_COL, rowHeader, "Interest Calculation Period*");
            writeString(ARREARS_TOLERANCE_COL, rowHeader, "Arrears Tolerance");
            writeString(REPAYMENT_STRATEGY_COL, rowHeader, "Repayment Strategy*");
            writeString(GRACE_ON_PRINCIPAL_PAYMENT_COL, rowHeader, "Grace-Principal Payment");
            writeString(GRACE_ON_INTEREST_PAYMENT_COL, rowHeader, "Grace-Interest Payment");
            writeString(GRACE_ON_INTEREST_CHARGED_COL, rowHeader, "Interest-Free Period(s)");
            writeString(INTEREST_CHARGED_FROM_COL, rowHeader, "Interest Charged From");
            writeString(FIRST_REPAYMENT_COL, rowHeader, "First Repayment On");
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
    			Row row = loanSheet.getRow(++rowIndex);
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
	    		//TODO: Clean this.
	    		CellRangeAddressList officeNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), OFFICE_NAME_COL, OFFICE_NAME_COL);
	        	CellRangeAddressList clientNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), CLIENT_NAME_COL, CLIENT_NAME_COL);
	        	CellRangeAddressList productNameRange = new  CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PRODUCT_COL, PRODUCT_COL);
	        	CellRangeAddressList loanOfficerRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_OFFICER_NAME_COL, LOAN_OFFICER_NAME_COL);
	        	CellRangeAddressList dateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), SUBMITTED_ON_DATE_COL, SUBMITTED_ON_DATE_COL);
	        	CellRangeAddressList fundNameRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), FUND_NAME_COL, FUND_NAME_COL);
	        	CellRangeAddressList principalRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), PRINCIPAL_COL, PRINCIPAL_COL);
	        	CellRangeAddressList noOfRepaymentsRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), NO_OF_REPAYMENTS_COL, NO_OF_REPAYMENTS_COL);
	        	CellRangeAddressList repaidFrequencyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAID_EVERY_FREQUENCY_COL, REPAID_EVERY_FREQUENCY_COL);
	        	CellRangeAddressList loanTermRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_TERM_COL, LOAN_TERM_COL);
	        	CellRangeAddressList loanTermFrequencyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), LOAN_TERM_FREQUENCY_COL, LOAN_TERM_FREQUENCY_COL);
	        	CellRangeAddressList disbursedDateRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), DISBURSED_DATE_COL, DISBURSED_DATE_COL);
	        	CellRangeAddressList interestFrequencyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), NOMINAL_INTEREST_RATE_FREQUENCY_COL, NOMINAL_INTEREST_RATE_FREQUENCY_COL);
	        	CellRangeAddressList interestRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), NOMINAL_INTEREST_RATE_COL, NOMINAL_INTEREST_RATE_COL);
	        	CellRangeAddressList amortizationRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), AMORTIZATION_COL, AMORTIZATION_COL);
	        	CellRangeAddressList interestMethodRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), INTEREST_METHOD_COL, INTEREST_METHOD_COL);
	        	CellRangeAddressList intrestCalculationPeriodRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), INTEREST_CALCULATION_PERIOD_COL, INTEREST_CALCULATION_PERIOD_COL);
	        	CellRangeAddressList repaymentStrategyRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), REPAYMENT_STRATEGY_COL, REPAYMENT_STRATEGY_COL);
	        	CellRangeAddressList arrearsToleranceRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), ARREARS_TOLERANCE_COL, ARREARS_TOLERANCE_COL);
	        	CellRangeAddressList graceOnPrincipalPaymentRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), GRACE_ON_PRINCIPAL_PAYMENT_COL, GRACE_ON_PRINCIPAL_PAYMENT_COL);
	        	CellRangeAddressList graceOnInterestPaymentRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), GRACE_ON_INTEREST_PAYMENT_COL, GRACE_ON_INTEREST_PAYMENT_COL);
	        	CellRangeAddressList graceOnInterestChargedRange = new CellRangeAddressList(1, SpreadsheetVersion.EXCEL97.getLastRowIndex(), GRACE_ON_INTEREST_CHARGED_COL, GRACE_ON_INTEREST_CHARGED_COL);
	        	
	        	DataValidationHelper validationHelper = new HSSFDataValidationHelper((HSSFSheet)worksheet);
	        	Workbook loanWorkbook = worksheet.getWorkbook();
	        	
	        	//Office Name
	        	Name officeGroup = loanWorkbook.createName();
	        	officeGroup.setNameName("Office");
	        	officeGroup.setRefersToFormula("Clients!$A$2:$A$" + (clientSheetPopulator.officeNames.size() + 1));
	        	
	        	ArrayList<String> officeNames = clientSheetPopulator.officeNames;
	        	List<Product> products = productSheetPopulator.getProducts();
	        	
	        	//Client Name
	        	Name[] clientGroups = new Name[officeNames.size()];
	        	ArrayList<String> formulas = new ArrayList<String>();
	        	for(Integer i = 0, j = 2; i < officeNames.size(); i++, j = j + 2) {
	        		String lastColumnLetters = CellReference.convertNumToColString(clientSheetPopulator.lastColumnLetters.get(i));
	        		formulas.add("Clients!$B$" + j + ":$" + lastColumnLetters + "$" + j);
	        		clientGroups[i] = loanWorkbook.createName();
	        	    clientGroups[i].setNameName(officeNames.get(i));
	        	    clientGroups[i].setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Product Name
	        	Name productGroup = loanWorkbook.createName();
	        	productGroup.setNameName("Products");
	        	productGroup.setRefersToFormula("Products!$B$2:$B$" + (productSheetPopulator.getProductsSize() + 1));
	        	
	        	//Loan Officer Name
	        	Name[] loanOfficerGroups = new Name[officeNames.size()];
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0, j = 2; i < officeNames.size(); i++, j = j + 2) {
	        		String lastColumnLetters = CellReference.convertNumToColString(personnelSheetPopulator.lastColumnLetters.get(i));
	        		formulas.add("Staff!$B$" + j + ":$" + lastColumnLetters + "$" + j);
	        		loanOfficerGroups[i] = loanWorkbook.createName();
	        	    loanOfficerGroups[i].setNameName(officeNames.get(i)+"X");
	        	    loanOfficerGroups[i].setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Fund Name
	        	Name fundGroup = loanWorkbook.createName();
	        	fundGroup.setNameName("Funds");
	        	fundGroup.setRefersToFormula("Funds!$B$2:$B$" + (fundSheetPopulator.getFundsSize() + 1));
	        	
	        	//Default Fund Names
	        	ArrayList<Name> fundOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$C$" + (i + 2));
	        		fundOfProduct.add(loanWorkbook.createName());
	        	    fundOfProduct.get(i).setNameName(products.get(i).getName() + "_Fund");
	        	    if(products.get(i).getFundName() != null) {
	        	       fundOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	    }
	        	}
	        	
	        	//Default Principal Names
	        	ArrayList<Name> principalOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$D$" + (i + 2));
	        		principalOfProduct.add(loanWorkbook.createName());
	        	    principalOfProduct.get(i).setNameName(products.get(i).getName() + "_Principal");
	        	    principalOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Min Principal Names
	        	ArrayList<Name> minPrincipalOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$E$" + (i + 2));
	        		minPrincipalOfProduct.add(loanWorkbook.createName());
	        	    minPrincipalOfProduct.get(i).setNameName(products.get(i).getName() + "_Min_Principal");
	        	    minPrincipalOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Max Principal Names
	        	ArrayList<Name> maxPrincipalOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$F$" + (i + 2));
	        		maxPrincipalOfProduct.add(loanWorkbook.createName());
	        	    maxPrincipalOfProduct.get(i).setNameName(products.get(i).getName() + "_Max_Principal");
	        	    maxPrincipalOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Default No. of Repayments Names
	        	ArrayList<Name> noOfRepaymentsOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$G$" + (i + 2));
	        		noOfRepaymentsOfProduct.add(loanWorkbook.createName());
	        		noOfRepaymentsOfProduct.get(i).setNameName(products.get(i).getName() + "_No_Repayment");
	        		noOfRepaymentsOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Default Min No. of Repayments Names
	        	ArrayList<Name> minRepaymentsOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$H$" + (i + 2));
	        		minRepaymentsOfProduct.add(loanWorkbook.createName());
	        		minRepaymentsOfProduct.get(i).setNameName(products.get(i).getName() + "_Min_Repayment");
	        		minRepaymentsOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Default Max No. of Repayments Names
	        	ArrayList<Name> maxRepaymentsOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$I$" + (i + 2));
	        		maxRepaymentsOfProduct.add(loanWorkbook.createName());
	        		maxRepaymentsOfProduct.get(i).setNameName(products.get(i).getName() + "_Max_Repayment");
	        		maxRepaymentsOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Default Repayment Every Name
	        	ArrayList<Name> repaymentEveryOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$J$" + (i + 2));
	        		repaymentEveryOfProduct.add(loanWorkbook.createName());
	        		repaymentEveryOfProduct.get(i).setNameName(products.get(i).getName() + "_Repayment_Every");
	        		repaymentEveryOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Default Repayment Every Frequency Name
	        	ArrayList<Name> repaymentFrequencyOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$K$" + (i + 2));
	        		repaymentFrequencyOfProduct.add(loanWorkbook.createName());
	        		repaymentFrequencyOfProduct.get(i).setNameName(products.get(i).getName() + "_REPAYMENT_FREQUENCY");
	        		repaymentFrequencyOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Default Interest Rate Names
	        	ArrayList<Name> interestOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$L$" + (i + 2));
	        		interestOfProduct.add(loanWorkbook.createName());
	        		interestOfProduct.get(i).setNameName(products.get(i).getName() + "_INTEREST");
	        		interestOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Min Interest Rate Names
	        	ArrayList<Name> minInterestOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$M$" + (i + 2));
	        		minInterestOfProduct.add(loanWorkbook.createName());
	        		minInterestOfProduct.get(i).setNameName(products.get(i).getName() + "_MIN_INTEREST");
	        		minInterestOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Max Interest Rate Names
	        	ArrayList<Name> maxInterestOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$N$" + (i + 2));
	        		maxInterestOfProduct.add(loanWorkbook.createName());
	        		maxInterestOfProduct.get(i).setNameName(products.get(i).getName() + "_MAX_INTEREST");
	        		maxInterestOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	

	        	//Fixed Interest Frequency Name
	        	ArrayList<Name> interestFrequencyOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$O$" + (i + 2));
	        		interestFrequencyOfProduct.add(loanWorkbook.createName());
	        		interestFrequencyOfProduct.get(i).setNameName(products.get(i).getName() + "_INTEREST_FREQUENCY");
	        		interestFrequencyOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Amortization Names
	        	ArrayList<Name> AmortizationOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$P$" + (i + 2));
	        		AmortizationOfProduct.add(loanWorkbook.createName());
	        		AmortizationOfProduct.get(i).setNameName(products.get(i).getName() + "_AMORTIZATION");
	        		AmortizationOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Interest Type Names
	        	ArrayList<Name> InterestTypeOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$Q$" + (i + 2));
	        		InterestTypeOfProduct.add(loanWorkbook.createName());
	        		InterestTypeOfProduct.get(i).setNameName(products.get(i).getName() + "_INTEREST_TYPE");
	        		InterestTypeOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Interest Calculation Period Names
	        	ArrayList<Name> InterestCalculationPeriodOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$R$" + (i + 2));
	        		InterestCalculationPeriodOfProduct.add(loanWorkbook.createName());
	        		InterestCalculationPeriodOfProduct.get(i).setNameName(products.get(i).getName() + "_INTEREST_CALCULATION");
	        		InterestCalculationPeriodOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Transaction Processing Strategy Names
	        	ArrayList<Name> TransactionProcessingStrategyOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$T$" + (i + 2));
	        		TransactionProcessingStrategyOfProduct.add(loanWorkbook.createName());
	        		TransactionProcessingStrategyOfProduct.get(i).setNameName(products.get(i).getName() + "_STRATEGY");
	        		TransactionProcessingStrategyOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Arrears Tolerance Names
	        	ArrayList<Name> ArrearsToleranceOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$S$" + (i + 2));
	        		ArrearsToleranceOfProduct.add(loanWorkbook.createName());
	        		ArrearsToleranceOfProduct.get(i).setNameName(products.get(i).getName() + "_ARREARS_TOLERANCE");
	        		ArrearsToleranceOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Grace On Principal Payment Names
	        	ArrayList<Name> GraceOnPrincipalPaymentOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$U$" + (i + 2));
	        		GraceOnPrincipalPaymentOfProduct.add(loanWorkbook.createName());
	        		GraceOnPrincipalPaymentOfProduct.get(i).setNameName(products.get(i).getName() + "_GRACE_PRINCIPAL");
	        		GraceOnPrincipalPaymentOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Grace On Interest Payment Names
	        	ArrayList<Name> GraceOnInterestPaymentOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$V$" + (i + 2));
	        		GraceOnInterestPaymentOfProduct.add(loanWorkbook.createName());
	        		GraceOnInterestPaymentOfProduct.get(i).setNameName(products.get(i).getName() + "_GRACE_INTEREST_PAYMENT");
	        		GraceOnInterestPaymentOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	//Grace On Principal Payment Names
	        	ArrayList<Name> GraceOnInterestChargedOfProduct = new ArrayList<Name>();
	        	formulas = new ArrayList<String>();
	        	for(Integer i = 0; i < products.size(); i++) {
	        		formulas.add("Products!$W$" + (i + 2));
	        		GraceOnInterestChargedOfProduct.add(loanWorkbook.createName());
	        		GraceOnInterestChargedOfProduct.get(i).setNameName(products.get(i).getName() + "_GRACE_INTEREST_CHARGED");
	        		GraceOnInterestChargedOfProduct.get(i).setRefersToFormula(formulas.get(i));
	        	}
	        	
	        	
	        	DataValidationConstraint officeNameConstraint = validationHelper.createExplicitListConstraint(clientSheetPopulator.getOfficeNames());
	        	DataValidationConstraint clientNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT($A1)");
	        	DataValidationConstraint productNameConstraint = validationHelper.createFormulaListConstraint("Products");
	        	DataValidationConstraint loanOfficerNameConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE($A1,\"X\"))");
	        	DataValidationConstraint activationDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=VLOOKUP($B1,$AA$2:$AB$" + (clientSheetPopulator.getClientsSize() + 1) + ",2,FALSE)", "=TODAY()", "dd/mm/yy");
	        	DataValidationConstraint disbursedDateConstraint = validationHelper.createDateConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=$E1", "=TODAY()", "dd/mm/yy");
	        	DataValidationConstraint fundNameConstraint = validationHelper.createFormulaListConstraint("Funds");
	        	DataValidationConstraint principalConstraint = validationHelper.createDecimalConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=INDIRECT(CONCATENATE($C1,\"_Min_Principal\"))", "=INDIRECT(CONCATENATE($C1,\"_Max_Principal\"))");
	        	DataValidationConstraint noOfRepaymentsConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=INDIRECT(CONCATENATE($C1,\"_Min_Repayment\"))", "=INDIRECT(CONCATENATE($C1,\"_Max_Repayment\"))");
	        	DataValidationConstraint frequencyConstraint = validationHelper.createExplicitListConstraint(new String[] {"Days","Weeks","Months"});
	        	DataValidationConstraint loanTermConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "=$H1*$I1", null);
	        	DataValidationConstraint interestFrequencyConstraint = validationHelper.createFormulaListConstraint("INDIRECT(CONCATENATE($C1,\"_INTEREST_FREQUENCY\"))");
	        	DataValidationConstraint interestConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.BETWEEN, "=INDIRECT(CONCATENATE($C1,\"_MIN_INTEREST\"))", "=INDIRECT(CONCATENATE($C1,\"_MAX_INTEREST\"))");
	        	DataValidationConstraint amortizationConstraint = validationHelper.createExplicitListConstraint(new String[] {"Equal principal payments","Equal Installments"});
	        	DataValidationConstraint interestMethodConstraint = validationHelper.createExplicitListConstraint(new String[] {"Flat","Declining Balance"});
	        	DataValidationConstraint interestCalculationPeriodConstraint = validationHelper.createExplicitListConstraint(new String[] {"Daily","Same as repayment period"});
	        	DataValidationConstraint repaymentStrategyConstraint = validationHelper.createExplicitListConstraint(new String[] {"Mifos style","Heavensfamily","Creocore","RBI (India)","Principal Interest Penalties Fees Order","Interest Principal Penalties Fees Order"});
	        	DataValidationConstraint arrearsToleranceConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	DataValidationConstraint graceOnPrincipalPaymentConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	DataValidationConstraint graceOnInterestPaymentConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	DataValidationConstraint graceOnInterestChargedConstraint = validationHelper.createIntegerConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null);
	        	
	        	DataValidation officeValidation = validationHelper.createValidation(officeNameConstraint, officeNameRange);
	        	officeValidation.setSuppressDropDownArrow(false);
	        	DataValidation clientValidation = validationHelper.createValidation(clientNameConstraint, clientNameRange);
	        	clientValidation.setSuppressDropDownArrow(false);
	        	DataValidation productNameValidation = validationHelper.createValidation(productNameConstraint, productNameRange);
	        	productNameValidation.setSuppressDropDownArrow(false);
	        	DataValidation loanOfficerValidation = validationHelper.createValidation(loanOfficerNameConstraint, loanOfficerRange);
	        	loanOfficerValidation.setSuppressDropDownArrow(false);
	        	DataValidation fundNameValidation = validationHelper.createValidation(fundNameConstraint, fundNameRange);
	        	fundNameValidation.setSuppressDropDownArrow(false);
	        	DataValidation repaidFrequencyValidation = validationHelper.createValidation(frequencyConstraint, repaidFrequencyRange);
	        	repaidFrequencyValidation.setSuppressDropDownArrow(false);
	        	DataValidation loanTermFrequencyValidation = validationHelper.createValidation(frequencyConstraint, loanTermFrequencyRange);
	        	loanTermFrequencyValidation.setSuppressDropDownArrow(false);
	        	DataValidation amortizationValidation = validationHelper.createValidation(amortizationConstraint, amortizationRange);
	        	amortizationValidation.setSuppressDropDownArrow(false);
	        	DataValidation interestMethodValidation = validationHelper.createValidation(interestMethodConstraint, interestMethodRange);
	        	interestMethodValidation.setSuppressDropDownArrow(false);
	        	DataValidation interestCalculationPeriodValidation = validationHelper.createValidation(interestCalculationPeriodConstraint, intrestCalculationPeriodRange);
	        	interestCalculationPeriodValidation.setSuppressDropDownArrow(false);
	        	DataValidation repaymentStrategyValidation = validationHelper.createValidation(repaymentStrategyConstraint, repaymentStrategyRange);
	        	repaymentStrategyValidation.setSuppressDropDownArrow(false);
	        	DataValidation activationDateValidation = validationHelper.createValidation(activationDateConstraint, dateRange);
	        	DataValidation disbursedDateValidation = validationHelper.createValidation(disbursedDateConstraint, disbursedDateRange);
	        	DataValidation principalValidation = validationHelper.createValidation(principalConstraint, principalRange);
	        	DataValidation loanTermValidation = validationHelper.createValidation(loanTermConstraint, loanTermRange);
	        	DataValidation noOfRepaymentsValidation = validationHelper.createValidation(noOfRepaymentsConstraint, noOfRepaymentsRange);
	        	DataValidation interestValidation = validationHelper.createValidation(interestConstraint, interestRange);
	        	DataValidation arrearsToleranceValidation = validationHelper.createValidation(arrearsToleranceConstraint, arrearsToleranceRange);
	        	DataValidation graceOnPrincipalPaymentValidation = validationHelper.createValidation(graceOnPrincipalPaymentConstraint, graceOnPrincipalPaymentRange);
	        	DataValidation graceOnInterestPaymentValidation = validationHelper.createValidation(graceOnInterestPaymentConstraint, graceOnInterestPaymentRange);
	        	DataValidation graceOnInterestChargedValidation = validationHelper.createValidation(graceOnInterestChargedConstraint, graceOnInterestChargedRange);
	        	DataValidation interestFrequencyValidation = validationHelper.createValidation(interestFrequencyConstraint, interestFrequencyRange);
	        	interestFrequencyValidation.setSuppressDropDownArrow(true);
	        	
	        	
	        	
	        	worksheet.addValidationData(officeValidation);
	            worksheet.addValidationData(clientValidation);
	            worksheet.addValidationData(productNameValidation);
	            worksheet.addValidationData(loanOfficerValidation);
	            worksheet.addValidationData(activationDateValidation);
	            worksheet.addValidationData(disbursedDateValidation);
	            worksheet.addValidationData(fundNameValidation);
	            worksheet.addValidationData(principalValidation);
	            worksheet.addValidationData(repaidFrequencyValidation);
	            worksheet.addValidationData(loanTermFrequencyValidation);
	            worksheet.addValidationData(noOfRepaymentsValidation);
	            worksheet.addValidationData(loanTermValidation);
	            worksheet.addValidationData(interestValidation);
	            worksheet.addValidationData(interestFrequencyValidation);
	            worksheet.addValidationData(amortizationValidation);
	            worksheet.addValidationData(interestMethodValidation);
	            worksheet.addValidationData(interestCalculationPeriodValidation);
	            worksheet.addValidationData(repaymentStrategyValidation);
	            worksheet.addValidationData(arrearsToleranceValidation);
	            worksheet.addValidationData(graceOnPrincipalPaymentValidation);
	            worksheet.addValidationData(graceOnInterestPaymentValidation);
	            worksheet.addValidationData(graceOnInterestChargedValidation);
	    	} catch (Exception e) {
	    		result.addError(e.getMessage());
	    	}
	       return result;
	    }
	    
	    private Result setDefaults(Sheet worksheet) {
	    	Result result = new Result();
	    	try {
	    		for(Integer rowNo = 1; rowNo < 1000; rowNo++)
	    		{
	    			Row row = worksheet.createRow(rowNo);
	    			row.createCell(FUND_NAME_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_Fund\"))),\"\",INDIRECT(CONCATENATE($C"+ (rowNo + 1) + ",\"_Fund\")))");
	    			row.createCell(PRINCIPAL_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_Principal\"))),\"\",INDIRECT(CONCATENATE($C"+ (rowNo + 1) + ",\"_Principal\")))");
	    			row.createCell(REPAID_EVERY_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_REPAYMENT_EVERY\"))),\"\",INDIRECT(CONCATENATE($C"+ (rowNo + 1) + ",\"_REPAYMENT_EVERY\")))");
	    			row.createCell(REPAID_EVERY_FREQUENCY_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_REPAYMENT_FREQUENCY\"))),\"\",INDIRECT(CONCATENATE($C"+ (rowNo + 1) + ",\"_REPAYMENT_FREQUENCY\")))");
	    			row.createCell(NO_OF_REPAYMENTS_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_No_Repayment\"))),\"\",INDIRECT(CONCATENATE($C"+ (rowNo + 1) + ",\"_No_Repayment\")))");
	    			row.createCell(LOAN_TERM_COL).setCellFormula("IF(ISERROR($H" + (rowNo + 1) + "*$I" + (rowNo + 1) + "),\"\",$H" + (rowNo + 1) + "*$I" + (rowNo + 1) + ")");
	    			row.createCell(LOAN_TERM_FREQUENCY_COL).setCellFormula("$J" + (rowNo + 1));
	    			row.createCell(NOMINAL_INTEREST_RATE_FREQUENCY_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST_FREQUENCY\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST_FREQUENCY\")))");
	    			row.createCell(NOMINAL_INTEREST_RATE_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST\")))");
	    			row.createCell(AMORTIZATION_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_AMORTIZATION\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_AMORTIZATION\")))");
	    			row.createCell(INTEREST_METHOD_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST_TYPE\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST_TYPE\")))");
	    			row.createCell(INTEREST_CALCULATION_PERIOD_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST_CALCULATION\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_INTEREST_CALCULATION\")))");
	    			row.createCell(ARREARS_TOLERANCE_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_ARREARS_TOLERANCE\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_ARREARS_TOLERANCE\")))");
	    			row.createCell(REPAYMENT_STRATEGY_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_STRATEGY\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_STRATEGY\")))");
	    			row.createCell(GRACE_ON_PRINCIPAL_PAYMENT_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_GRACE_PRINCIPAL\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_GRACE_PRINCIPAL\")))");
	    			row.createCell(GRACE_ON_INTEREST_PAYMENT_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_GRACE_INTEREST_PAYMENT\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_GRACE_INTEREST_PAYMENT\")))");
	    			row.createCell(GRACE_ON_INTEREST_CHARGED_COL).setCellFormula("IF(ISERROR(INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_GRACE_INTEREST_CHARGED\"))),\"\",INDIRECT(CONCATENATE($C" + (rowNo + 1) + ",\"_GRACE_INTEREST_CHARGED\")))");
	    		}
	    		
	    	} catch (Exception e) {
	    		result.addError(e.getMessage());
	    	}
	       return result;
	    	}
}