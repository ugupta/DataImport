package org.openmf.mifos.dataimport;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openmf.mifos.dataimport.dto.AuthToken;

import com.google.gson.Gson;

public abstract class AbstractDataImportHandler implements DataImportHandler {

    protected final Sheet sheet;

    private final String baseURL;

    private final String authToken;

    private final String userName;

    private final String password;

    private final String tenantId;

    public AbstractDataImportHandler(Sheet sheet) {
        this.sheet = sheet;
        baseURL = "http://localhost:8080/api/v1/"; // System.getProperty("mifos.endpoint");
        userName = "mifos"; // System.getProperty("mifos.user.id");
        password = "password"; // System.getProperty("mifos.password");
        tenantId = "default"; // System.getProperty("mifos.tenant.id");
        authToken = createAuthToken();
    };

    private String createAuthToken() {
        String url = baseURL + "authentication?username=" + userName + "&password=" + password;
        try {

            SimpleHttpClient sClient = new SimpleHttpClient(url, SimpleHttpClient.Method.POST);
            sClient.header("X-Mifos-Platform-TenantId", tenantId).header("Content-Type", "application/json; charset=utf-8");
            sClient.execute();
            // might have to check IO close
            return new Gson().fromJson(sClient.getContent(), AuthToken.class).getBase64EncodedAuthenticationKey();
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    protected void post(String path, String payload) {
        String url = baseURL + path;

        try {

            SimpleHttpClient sClient = new SimpleHttpClient(url, SimpleHttpClient.Method.POST);
            sClient.header("Authorization", "Basic " + authToken);
            sClient.header("X-Mifos-Platform-TenantId", tenantId);
            sClient.header("Content-Type", "application/json; charset=utf-8");

            sClient.sendContent(payload);
            if (sClient.status() != HttpURLConnection.HTTP_OK) { throw new IllegalStateException("failed with status " + sClient.status()); }
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    protected Integer getNumberOfRows() {
        Integer noOfEntries = 1;
        // getLastRowNum and getPhysicalNumberOfRows showing false values
        // sometimes.
        while (sheet.getRow(noOfEntries) != null) {
            noOfEntries++;
        }
        return noOfEntries;
    }

    protected int readAsInt(int colIndex, Row row) {
        return ((Double) row.getCell(colIndex).getNumericCellValue()).intValue();
    }

    protected String readAsString(int colIndex, Row row) {
        return row.getCell(colIndex).getStringCellValue();
    }

    protected Date readAsDate(int colIndex, Row row) {
        return row.getCell(colIndex).getDateCellValue();
    }

}
