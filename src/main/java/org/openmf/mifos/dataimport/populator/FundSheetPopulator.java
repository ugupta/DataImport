package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Fund;
import org.openmf.mifos.dataimport.handler.Result;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class FundSheetPopulator extends AbstractWorkbookPopulator {

    private static final Logger logger = LoggerFactory.getLogger(FundSheetPopulator.class);
	
	private final RestClient client;
	
	private String content;
	
	private List<Fund> funds;
	
	public static final int ID_COL = 0;
	public static final int NAME_COL = 1;
	
	public FundSheetPopulator(RestClient client) {
        this.client = client;
    }
    
    @Override
    public Result downloadAndParse() {
    	Result result = new Result();
        try {
        	client.createAuthToken();
        	funds = new ArrayList<Fund>();
            content = client.get("funds");
            Gson gson = new Gson();
            JsonElement json = new JsonParser().parse(content);
            JsonArray array = json.getAsJsonArray();
            Iterator<JsonElement> iterator = array.iterator();
            while(iterator.hasNext()) {
            	json = iterator.next();
            	Fund fund = gson.fromJson(json, Fund.class);
            	funds.add(fund);
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
        Sheet fundSheet = workbook.createSheet("Funds");
        setLayout(fundSheet);
        for(Fund fund : funds) {
        	Row row = fundSheet.createRow(rowIndex++);
        	writeInt(ID_COL, row, fund.getId());
        	writeString(NAME_COL, row, fund.getName().trim().replaceAll("[ )(]", "_"));
         }
    	} catch (Exception e) {
    		result.addError(e.getMessage());
    		logger.error(e.getMessage());
    	}
        return result;
    }
    
    private void setLayout(Sheet worksheet) {
    	worksheet.setColumnWidth(ID_COL, 2000);
        worksheet.setColumnWidth(NAME_COL, 7000);
        Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        writeString(ID_COL, rowHeader, "ID");
        writeString(NAME_COL, rowHeader, "Name");
    }
    
    public Integer getFundsSize() {
		 return funds.size();
	 }
}
