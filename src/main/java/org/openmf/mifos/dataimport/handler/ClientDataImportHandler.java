package org.openmf.mifos.dataimport.handler;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Client;
import org.openmf.mifos.dataimport.dto.CorporateClient;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class ClientDataImportHandler extends AbstractDataImportHandler {
           
	private static final Logger logger = LoggerFactory.getLogger(ClientDataImportHandler.class);
	
	public static final int FIRST_NAME_COL = 0;
	public static final int FULL_NAME_COL = 0;
    public static final int LAST_NAME_COL = 1;
    public static final int MIDDLE_NAME_COL = 2;
    public static final int OFFICE_NAME_COL = 3;
    public static final int STAFF_NAME_COL = 4;
    public static final int EXTERNAL_ID_COL = 5;
    public static final int ACTIVATION_DATE_COL = 6;
    public static final int ACTIVE_COL = 7;
    public static final int STATUS_COL = 8;

    private List<Client> clients = new ArrayList<Client>();
    private List<CorporateClient> corporateClients = new ArrayList<CorporateClient>();
    private String clientType;
    
    private final RestClient restClient;
    
    private final Workbook workbook;

    public ClientDataImportHandler(Workbook workbook, RestClient client) {
        this.workbook = workbook;
        this.restClient = client;
    }
    
    @Override
    public Result parse() {
        Result result = new Result();
        Sheet clientSheet = workbook.getSheet("Clients");
        Integer noOfEntries = getNumberOfRows(clientSheet, 0);
        if(readAsString(FIRST_NAME_COL, clientSheet.getRow(0)).equals("First Name*"))
        	clientType = "Individual";
        else
        	clientType = "Corporate";
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = clientSheet.getRow(rowIndex);
                String status = readAsString(STATUS_COL, row);
                if(status.equals("Imported"))
                	continue;
                String officeName = readAsString(OFFICE_NAME_COL, row);
                String officeId = getIdByName(workbook.getSheet("Offices"), officeName).toString();
                String staffName = readAsString(STAFF_NAME_COL, row);
                String staffId = getIdByName(workbook.getSheet("Staff"), staffName).toString();
                String externalId = readAsString(EXTERNAL_ID_COL, row);
                String activationDate = readAsDate(ACTIVATION_DATE_COL, row);
                String active = readAsBoolean(ACTIVE_COL, row).toString();
                if(clientType.equals("Individual")) {
                    String firstName = readAsString(FIRST_NAME_COL, row);
                    String lastName = readAsString(LAST_NAME_COL, row);
                    String middleName = readAsString(MIDDLE_NAME_COL, row);
                    clients.add(new Client ( firstName, lastName, middleName, activationDate, active, externalId, officeId, staffId, rowIndex));
                  } else {
                    String fullName = readAsString(FULL_NAME_COL, row);
                    corporateClients.add(new CorporateClient(fullName, activationDate, active, externalId, officeId, staffId, rowIndex));
                  }
                
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
        Sheet clientSheet = workbook.getSheet("Clients");
        restClient.createAuthToken();
        if(clientType.equals("Individual"))
        for (Client client : clients) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(client);
                logger.info(payload);
                restClient.post("clients", payload);
                Cell statusCell = clientSheet.getRow(client.getRowIndex()).createCell(STATUS_COL);
                statusCell.setCellValue("Imported");
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
            } catch (Exception e) {
            	String message = parseStatus(e.getMessage());
            	Cell statusCell = clientSheet.getRow(client.getRowIndex()).createCell(STATUS_COL);
            	statusCell.setCellValue(message);
                statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
                result.addError("Row = " + client.getRowIndex() + " ," + message);
            }
        }
        else {
        	for (CorporateClient corporateClient : corporateClients) {
                try {
                    Gson gson = new Gson();
                    String payload = gson.toJson(corporateClient);
                    logger.info(payload);
                    restClient.post("clients", payload);
                    Cell statusCell = clientSheet.getRow(corporateClient.getRowIndex()).createCell(STATUS_COL);
                    statusCell.setCellValue("Imported");
                    statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.LIGHT_GREEN));
                } catch (Exception e) {
                	String message = parseStatus(e.getMessage());
                	Cell statusCell = clientSheet.getRow(corporateClient.getRowIndex()).createCell(STATUS_COL);
                	statusCell.setCellValue(message);
                    statusCell.setCellStyle(getCellStyle(workbook, IndexedColors.RED));
                    result.addError("Row = " + corporateClient.getRowIndex() + " ," + message);                
                    }
            }
        }
        clientSheet.setColumnWidth(STATUS_COL, 15000);
    	writeString(STATUS_COL, clientSheet.getRow(0), "Status");
        return result;
    }
    

    
    public List<Client> getClients() {
        return clients;
    }
    
    public List<CorporateClient> getCorporateClients() {
        return corporateClients;
    }
}
