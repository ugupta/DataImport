package org.openmf.mifos.dataimport;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Office;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class OfficeSheetPopulator extends AbstractWorkbookPopulator {
	
	private static final Logger logger = LoggerFactory.getLogger(OfficeSheetPopulator.class);
	
	private final RestClient client;
	
	private String content;
	
	private List<Office> offices = new ArrayList<Office>();

    public OfficeSheetPopulator(RestClient client) {
        this.client = client;
    }
    
    @Override
    public Result download() {
        Result result = new Result();
        client.createAuthToken();
        try {
            content = client.get("offices");
        } catch (Exception e) {
            result.addError(e.getMessage());
        }
        return result;
    }

    @Override
    public Result populate(Workbook workbook) {
        Result result = new Result();
        Sheet officeSheet = workbook.createSheet("Offices");
        Gson gson = new Gson();
//        content = "{\"finally\":"+content+"}";
        logger.info(content);
        JsonElement json = new JsonParser().parse(content);
        JsonArray array = json.getAsJsonArray();
        Iterator<JsonElement> iterator = array.iterator();
        while(iterator.hasNext()) {
        	JsonElement json2 = iterator.next();
        	Office office = gson.fromJson(json2, Office.class);
        	offices.add(office);
        	logger.info("CHECK : "+office.toString());
        }
        return result;
    }
    
    public List<Office> getOffices() {
        return offices;
    }

}
