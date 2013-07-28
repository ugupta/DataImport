package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Product;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class ProductSheetPopulator extends AbstractWorkbookPopulator {
	
	private static final Logger logger = LoggerFactory.getLogger(ProductSheetPopulator.class);
	
    private final RestClient client;
	
	private String content;
	
	private List<Product> products;
	
	public static final int ID_COL = 0;
	public static final int NAME_COL = 1;
	public static final int FUND_NAME_COL = 2;
	
	public ProductSheetPopulator(RestClient client) {
        this.client = client;
    }
	
	@Override
    public Result downloadAndParse() {
    	Result result = new Result();
        try {
        	client.createAuthToken();
        	products = new ArrayList<Product>();
            content = client.get("loanproducts");
            Gson gson = new Gson();
            JsonElement json = new JsonParser().parse(content);
            JsonArray array = json.getAsJsonArray();
            Iterator<JsonElement> iterator = array.iterator();
            while(iterator.hasNext()) {
            	json = iterator.next();
            	Product product = gson.fromJson(json, Product.class);
            	if(product.getStatus().equals("loanProduct.active"))
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
	            for(Product product : products) {
	            	Row row = productSheet.createRow(rowIndex++);
	            	writeInt(ID_COL, row, product.getId());
	            	writeString(NAME_COL, row, product.getName().trim().replaceAll("[ )(]", "_"));
	            	if(product.getFundName() != null)
	            	    writeString(FUND_NAME_COL, row, product.getFundName());
	            }
	    	} catch (Exception e) {
	    		result.addError(e.getMessage());
	    		logger.error(e.getMessage());
	    	}
	        return result;
	 }
	 
	 private void setLayout(Sheet worksheet) {
		    worksheet.setColumnWidth(ID_COL, 2000);
	        worksheet.setColumnWidth(NAME_COL, 5000);
	        worksheet.setColumnWidth(FUND_NAME_COL, 3000);
	        Row rowHeader = worksheet.createRow(0);
	        rowHeader.setHeight((short)500);
	        writeString(ID_COL, rowHeader, "ID");
	        writeString(NAME_COL, rowHeader, "Name");
	        writeString(FUND_NAME_COL, rowHeader, "Fund");
	 }
	 
	 public Integer getProductsSize() {
		 return products.size();
	 }
	 

}
