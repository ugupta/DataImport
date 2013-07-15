package org.openmf.mifos.dataimport.handler;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openmf.mifos.dataimport.dto.Client;
import org.openmf.mifos.dataimport.http.RestClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.gson.Gson;

public class ClientDataImportHandler extends AbstractDataImportHandler {
           
	private static final Logger logger = LoggerFactory.getLogger(ClientDataImportHandler.class);
	
	public static final int FIRST_NAME_COL = 0;
    public static final int LAST_NAME_COL = 1;
    public static final int MIDDLE_NAME_COL = 2;
    public static final int OFFICE_NAME_COL = 3;
    public static final int STAFF_NAME_COL = 4;
    public static final int EXTERNAL_ID_COL = 5;
    public static final int ACTIVATION_DATE_COL = 6;
    public static final int ACTIVE_COL = 7;

    private List<Client> clients = new ArrayList<Client>();
    
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
        Integer noOfEntries = getNumberOfRows(clientSheet);
        for (int rowIndex = 1; rowIndex < noOfEntries; rowIndex++) {
            Row row;
            try {
                row = clientSheet.getRow(rowIndex);
                String firstName = readAsString(FIRST_NAME_COL, row);
                String lastName = readAsString(LAST_NAME_COL, row);
                String middleName = readAsString(MIDDLE_NAME_COL, row);
                String officeName = readAsString(OFFICE_NAME_COL, row);
                String officeId = getIdByName(workbook.getSheet("Offices"), officeName).toString();
                String staffName = readAsString(STAFF_NAME_COL, row);
                String staffId = getIdByName(workbook.getSheet("Staff"), staffName).toString();
                String externalId = readAsString(EXTERNAL_ID_COL, row);
                String activationDate = readAsDate(ACTIVATION_DATE_COL, row);
                String active = readAsBoolean(ACTIVE_COL, row).toString();
                clients.add(new Client ( firstName, lastName, middleName, activationDate, active, externalId, officeId, staffId, rowIndex));
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
        restClient.createAuthToken();
        for (Client client : clients) {
            try {
                Gson gson = new Gson();
                String payload = gson.toJson(client);
                logger.info(payload);
                restClient.post("clients", payload);
            } catch (Exception e) {
                logger.error("row = " + client.getRowIndex(), e);
                result.addError("Row = " + client.getRowIndex() + " ," + e.getMessage());
            }

        }
        return result;
    }
    
    public List<Client> getClients() {
        return clients;
    }
}