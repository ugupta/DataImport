package org.openmf.mifos.dataimport;

import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Date;

import org.apache.http.HttpHeaders;
import org.apache.http.HttpResponse;
import org.apache.http.HttpStatus;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.DefaultHttpClient;
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
        HttpClient client = new DefaultHttpClient();
        HttpPost postRequest = new HttpPost(baseURL + "authentication?username=" + userName + "&password=" + password);
        postRequest.setHeader("X-Mifos-Platform-TenantId", tenantId);
        postRequest.setHeader(HttpHeaders.CONTENT_TYPE, "application/json; charset=utf-8");

        try {
            HttpResponse response = client.execute(postRequest);
            // might have to check IO close
            return new Gson().fromJson(new InputStreamReader(response.getEntity().getContent(), "UTF-8"), AuthToken.class).getBase64EncodedAuthenticationKey();
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    protected void post(String url, String payload) {
        HttpClient client = new DefaultHttpClient();
        HttpPost postRequest = new HttpPost(baseURL + url);
        postRequest.setHeader(HttpHeaders.AUTHORIZATION, "Basic " + authToken);
        postRequest.setHeader("X-Mifos-Platform-TenantId", tenantId);
        postRequest.setHeader(HttpHeaders.CONTENT_TYPE, "application/json; charset=utf-8");
        try {
            postRequest.setEntity(new StringEntity(payload, "UTF-8"));
            HttpResponse response = client.execute(postRequest);
            if (response.getStatusLine().getStatusCode() != HttpStatus.SC_OK) { throw new IllegalStateException("failed : " + response.getStatusLine()); }
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
