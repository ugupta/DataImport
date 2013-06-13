package org.openmf.mifos.dataimport;

import java.io.DataInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openmf.mifos.dataimport.dto.AuthToken;
import org.openmf.mifos.dataimport.http.HttpRequestBuilder;
import org.openmf.mifos.dataimport.http.SimpleHttpRequest.Header;
import org.openmf.mifos.dataimport.http.SimpleHttpRequest.Method;
import org.openmf.mifos.dataimport.http.SimpleHttpResponse;

import com.google.gson.Gson;

public abstract class AbstractDataImportHandler implements DataImportHandler {

    protected final Sheet sheet;

    private final String baseURL;

    private final String userName;

    private final String password;

    private final String tenantId;

    private String authToken;

    public AbstractDataImportHandler(Sheet sheet) {
        this.sheet = sheet;
        baseURL = "http://localhost:8080/api/v1/"; // System.getProperty("mifos.endpoint");
        userName = "mifos"; // System.getProperty("mifos.user.id");
        password = "password"; // System.getProperty("mifos.password");
        tenantId = "default"; // System.getProperty("mifos.tenant.id");
    };

    private String createAuthToken() {
        String url = baseURL + "authentication?username=" + userName + "&password=" + password;
        try {

            SimpleHttpResponse response = new HttpRequestBuilder().withURL(url).withMethod(Method.POST)
            		.addHeader(Header.MIFOS_TENANT_ID, tenantId)
            		.addHeader(Header.CONTENT_TYPE, "application/json; charset=utf-8").execute();

            // might have to check IO close
            return new Gson().fromJson(readContentAndClose(response.getContent()), AuthToken.class).getBase64EncodedAuthenticationKey();
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    private String readContentAndClose(InputStream content) throws IOException {
        DataInputStream stream = new DataInputStream(content);
        String data = stream.readUTF();
        stream.close();
        return data;
    }

    protected void post(String path, String payload) {
        authToken = createAuthToken();
        String url = baseURL + path;

        try {

        	SimpleHttpResponse response = new HttpRequestBuilder().withURL(url).withMethod(Method.POST)
        			.addHeader(Header.AUTHORIZATION, "Basic " + authToken)
    				.addHeader(Header.MIFOS_TENANT_ID, tenantId)
    				.addHeader(Header.CONTENT_TYPE, "application/json; charset=utf-8")
    				.withContent(payload).execute();

            if (response.getStatus() != HttpURLConnection.HTTP_OK) { throw new IllegalStateException("failed with status " + response.getStatus()); }
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
        try {
            return row.getCell(colIndex).getStringCellValue();
        } catch (Exception e) {
            return row.getCell(colIndex).getNumericCellValue() + "";
        }
    }

    protected Date readAsDate(int colIndex, Row row) {
        return row.getCell(colIndex).getDateCellValue();
    }

}
