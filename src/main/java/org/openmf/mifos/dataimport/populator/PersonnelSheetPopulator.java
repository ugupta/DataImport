package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Personnel;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

public class PersonnelSheetPopulator extends AbstractWorkbookPopulator {

    private static final Logger logger = LoggerFactory.getLogger(PersonnelSheetPopulator.class);
	
	private final RestClient client;
	
	private String content;
	
	private List<Personnel> personnel = new ArrayList<Personnel>();
	
	public static final int ID_COL = 0;
	public static final int FULL_NAME_COL = 1;
	public static final int OFFICE_ID_COL = 2;
	public static final int OFFICE_NAME_COL = 3;
	
	public PersonnelSheetPopulator(RestClient client) {
        this.client = client;
    }
	
	 @Override
	    public void downloadAndParse() {
	        client.createAuthToken();
	        try {
	            content = client.get("staff");
	            Gson gson = new Gson();
	            JsonElement json = new JsonParser().parse(content);
	            JsonArray array = json.getAsJsonArray();
	            Iterator<JsonElement> iterator = array.iterator();
	            while(iterator.hasNext()) {
	            	JsonElement json2 = iterator.next();
	            	Personnel person = gson.fromJson(json2, Personnel.class);
	            	personnel.add(person);
	            	logger.info("CHECK : " + person.toString());
	            }
	        } catch (Exception e) {
	            
	        }
	    }

	    @Override
	    public void populate(Workbook workbook) {
	        int rowIndex = 1;
	        Sheet staffSheet = workbook.createSheet("Staff");
	        setLayout(staffSheet);
	        for(Personnel person:personnel) {
	        	if(person.isLoanOfficer()) {
	        	Row row = staffSheet.createRow(rowIndex);
	        	writeInt(ID_COL, row, person.getId());
	        	writeString(FULL_NAME_COL, row, person.getFirstName() + " " + person.getLastName());
	        	writeInt(OFFICE_ID_COL, row, person.getOfficeId());
	        	writeString(OFFICE_NAME_COL, row, person.getOfficeName());
	        	rowIndex++;
	        	}
	        }
	        staffSheet.protectSheet("");
	    }
	    
	    public void setLayout(Sheet worksheet) {
	    	worksheet.setColumnWidth(0, 3000);
	        worksheet.setColumnWidth(1, 7000);
	        worksheet.setColumnWidth(2, 3000);
	        worksheet.setColumnWidth(3, 7000);
	        Row rowHeader = worksheet.createRow(0);
	        rowHeader.setHeight((short)500);
	        writeString(ID_COL,rowHeader, "ID");
	        writeString(FULL_NAME_COL, rowHeader, "Name");
	        writeString(OFFICE_ID_COL, rowHeader, "Office ID");
	        writeString(OFFICE_NAME_COL, rowHeader, "Office Name");
	    }
	    
	    public List<Personnel> getPersonnel() {
	        return personnel;
	    }
	    
	    public Integer getPersonnelSize() {
	    	return personnel.size();
	    }
	    
	    public String[] getPersonnelNames() {
	    	List<String> personnelNameList = new ArrayList<String>();
	    	for (int i = 0; i < personnel.size(); i++)
	    		if(personnel.get(i).isLoanOfficer())
	    		  personnelNameList.add(personnel.get(i).getFirstName() + " " + personnel.get(i).getLastName()); 
	    	return personnelNameList.toArray(new String[personnelNameList.size()]);
	    }
}
