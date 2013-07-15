package org.openmf.mifos.dataimport.populator;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
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
	
	public static final int ID_COL = 0;
	public static final int OFFICE_NAME_COL = 1;
	public static final int EXTERNAL_ID_COL = 2;
	public static final int OPENING_DATE_COL = 3;
    public static final int PARENT_NAME_COL = 4;
    public static final int HIERARCHY_COL = 5;

    public OfficeSheetPopulator(RestClient client) {
        this.client = client;
    }
    
    @Override
    public void downloadAndParse() {
        client.createAuthToken();
        try {
            content = client.get("offices");
            Gson gson = new Gson();
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
        } catch (Exception e) {
            
        }
    }

    @Override
    public void populate(Workbook workbook) {
        int rowIndex = 1;
        Sheet officeSheet = workbook.createSheet("Offices");
        setLayout(officeSheet);
        CellStyle dateCellStyle = workbook.createCellStyle();
        short df = workbook.createDataFormat().getFormat("dd/mm/yy");
        dateCellStyle.setDataFormat(df);
        for(Office office:offices) {
        	Row row = officeSheet.createRow(rowIndex);
        	writeInt(ID_COL, row, office.getId());
        	writeString(OFFICE_NAME_COL, row, office.getName());
        	writeString(EXTERNAL_ID_COL, row, office.getExternalId());
        	writeDate(OPENING_DATE_COL, row, ""+office.getOpeningDate().get(1)+"/"+office.getOpeningDate().get(2)+"/"+office.getOpeningDate().get(0), dateCellStyle);
        	writeString(PARENT_NAME_COL,row,office.getParentName());
        	writeString(HIERARCHY_COL, row, office.getHierarchy());
        	rowIndex++;
        }
        officeSheet.protectSheet("");
    }
    
    public void setLayout(Sheet worksheet) {
    	worksheet.setColumnWidth(0, 2000);
        worksheet.setColumnWidth(1, 7000);
        worksheet.setColumnWidth(2, 7000);
        worksheet.setColumnWidth(3, 4000);
        worksheet.setColumnWidth(4, 7000);
        worksheet.setColumnWidth(5, 6000);
        Row rowHeader = worksheet.createRow(0);
        rowHeader.setHeight((short)500);
        writeString(ID_COL, rowHeader, "ID");
        writeString(OFFICE_NAME_COL, rowHeader, "Name");
        writeString(EXTERNAL_ID_COL, rowHeader, "External ID");
        writeString(OPENING_DATE_COL, rowHeader, "Opening Date");
        writeString(PARENT_NAME_COL, rowHeader, "Parent Name");
        writeString(HIERARCHY_COL, rowHeader, "Hierarchy");
    }
    
    public List<Office> getOffices() {
        return offices;
    }
    
    public Integer getOfficeSize() {
    	return offices.size();
    }
    
    public String[] getOfficeNames() {
    	String[] officeNameList = new String[offices.size()];
    	for (int i = 0; i < offices.size(); i++)
    		officeNameList[i] = offices.get(i).getName();
    	return officeNameList;
    }

}
