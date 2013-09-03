package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.SavingsProduct;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class SavingsProductSheetPopulator extends AbstractWorkbookPopulator {
	
private static final Logger logger = LoggerFactory.getLogger(SavingsProductSheetPopulator.class);
	
    private final RestClient client;
	
	private String content;
	
	private static final int ID_COL = 0;
	private static final int NAME_COL = 1;
	private static final int NOMINAL_ANNUAL_INTEREST_RATE_COL = 2;
	private static final int INTEREST_COMPOUNDING_PERIOD_COL = 3;
	private static final int INTEREST_POSTING_PERIOD_COL = 4;
	private static final int INTEREST_CALCULATION_COL = 5;
	private static final int INTEREST_CALCULATION_DAYS_IN_YEAR_COL = 6;
	private static final int MIN_OPENING_BALANCE_COL = 7;
	private static final int LOCKIN_PERIOD_COL = 8;
	private static final int LOCKIN_PERIOD_FREQUENCY_COL = 9;
	private static final int WITHDRAWAL_FEE_AMOUNT_COL = 10;
	private static final int WITHDRAWAL_FEE_TYPE_COL = 11;
	private static final int ANNUAL_FEE_COL = 12;
	private static final int ANNUAL_FEE_ON_MONTH_DAY_COL = 13;
	
	private List<SavingsProduct> products;
	
	public SavingsProductSheetPopulator(RestClient client) {
        this.client = client;
    }
	
	@Override
    public Result downloadAndParse() {
    	Result result = new Result();
        try {
        	client.createAuthToken();
        	products = new ArrayList<SavingsProduct>();
            content = client.get("savingsproducts");
            Gson gson = new Gson();
            JsonElement json = new JsonParser().parse(content);
            JsonArray array = json.getAsJsonArray();
            Iterator<JsonElement> iterator = array.iterator();
            while(iterator.hasNext()) {
            	json = iterator.next();
            	SavingsProduct product = gson.fromJson(json, SavingsProduct.class);
            	products.add(product);
            }
        } catch (Exception e) {
            result.addError(e.getMessage());
            logger.error(e.getMessage());
        }
        return result;
    }

	@Override
	 public Result populate(Workbook workbook) {
	    	Result result = new Result();
	    	try{
	    		int rowIndex = 1;
	            Sheet productSheet = workbook.createSheet("Products");
	            setLayout(productSheet);
	            CellStyle dateCellStyle = workbook.createCellStyle();
	            short df = workbook.createDataFormat().getFormat("dd-mmm");
	            dateCellStyle.setDataFormat(df);
	            for(SavingsProduct product : products) {
	            	Row row = productSheet.createRow(rowIndex++);
	            	writeInt(ID_COL, row, product.getId());
	            	writeString(NAME_COL, row, product.getName().trim().replaceAll("[ )(]", "_"));
	            	writeDouble(NOMINAL_ANNUAL_INTEREST_RATE_COL, row, product.getNominalAnnualInterestRat());
	            	writeString(INTEREST_COMPOUNDING_PERIOD_COL, row, product.getInterestCompoundingPeriodType().getValue());
	            	writeString(INTEREST_POSTING_PERIOD_COL, row, product.getInterestPostingPeriodType().getValue());
	            	writeString(INTEREST_CALCULATION_COL, row, product.getInterestCalculationType().getValue());
	            	writeString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, row, product.getInterestCalculationDaysInYearType().getValue());
	            	writeDouble(MIN_OPENING_BALANCE_COL, row, product.getMinRequiredOpeningBalance());
	            	writeInt(LOCKIN_PERIOD_COL, row, product.getLockinPeriodFrequency());
	            	writeString(LOCKIN_PERIOD_FREQUENCY_COL, row, product.getLockinPeriodFrequencyType().getValue());
	            	writeDouble(WITHDRAWAL_FEE_AMOUNT_COL, row, product.getWithdrawalFeeAmount());
	            	writeString(WITHDRAWAL_FEE_TYPE_COL, row, product.getWithdrawalFeeType().getValue());
	            	writeDouble(ANNUAL_FEE_COL, row, product.getAnnualFeeAmount());
	            	writeDate(ANNUAL_FEE_ON_MONTH_DAY_COL, row, product.getAnnualFeeOnMonthDay().get(1) + "/" + product.getAnnualFeeOnMonthDay().get(0) + "/2010" , dateCellStyle);
	            }
	            
	        	productSheet.protectSheet("");
    	} catch (Exception e) {
    		result.addError(e.getMessage());
    		logger.error(e.getMessage());
    	}
        return result;
      }
	
	private void setLayout(Sheet worksheet) {
		Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        worksheet.setColumnWidth(ID_COL, 2000);
        worksheet.setColumnWidth(NAME_COL, 5000);
        worksheet.setColumnWidth(NOMINAL_ANNUAL_INTEREST_RATE_COL, 2000);
        worksheet.setColumnWidth(INTEREST_COMPOUNDING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_POSTING_PERIOD_COL, 3000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_COL, 3000);
        worksheet.setColumnWidth(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, 3000);
        worksheet.setColumnWidth(MIN_OPENING_BALANCE_COL, 3000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_COL, 3000);
        worksheet.setColumnWidth(LOCKIN_PERIOD_FREQUENCY_COL, 3000);
        worksheet.setColumnWidth(WITHDRAWAL_FEE_AMOUNT_COL, 3000);
        worksheet.setColumnWidth(WITHDRAWAL_FEE_TYPE_COL, 3000);
        worksheet.setColumnWidth(ANNUAL_FEE_COL, 3000);
        worksheet.setColumnWidth(ANNUAL_FEE_ON_MONTH_DAY_COL, 3000);
        
        writeString(ID_COL, rowHeader, "ID");
        writeString(NAME_COL, rowHeader, "Name");
        writeString(NOMINAL_ANNUAL_INTEREST_RATE_COL, rowHeader, "Interest");
        writeString(INTEREST_COMPOUNDING_PERIOD_COL, rowHeader, "Interest Compounding Period");
        writeString(INTEREST_POSTING_PERIOD_COL, rowHeader, "Interest Posting Period");
        writeString(INTEREST_CALCULATION_COL, rowHeader, "Interest Calculated Using");
        writeString(INTEREST_CALCULATION_DAYS_IN_YEAR_COL, rowHeader, "# Days In Year");
        writeString(MIN_OPENING_BALANCE_COL, rowHeader, "Min Opening Balance");
        writeString(LOCKIN_PERIOD_COL, rowHeader, "Locked In For");
        writeString(LOCKIN_PERIOD_FREQUENCY_COL, rowHeader, "Frequency");
        writeString(WITHDRAWAL_FEE_AMOUNT_COL, rowHeader, "Withdrawal Fee");
        writeString(WITHDRAWAL_FEE_TYPE_COL, rowHeader, "Type");
        writeString(ANNUAL_FEE_COL, rowHeader, "Annual Fee");
        writeString(ANNUAL_FEE_ON_MONTH_DAY_COL, rowHeader, "On");
	}
	
	public List<SavingsProduct> getProducts() {
		 return products;
	 }
	
	public Integer getProductsSize() {
		 return products.size();
	 }
}
