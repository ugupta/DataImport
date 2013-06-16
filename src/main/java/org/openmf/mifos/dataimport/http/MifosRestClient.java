package org.openmf.mifos.dataimport.http;

import java.io.DataInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;

import org.openmf.mifos.dataimport.dto.AuthToken;
import org.openmf.mifos.dataimport.http.SimpleHttpRequest.Method;

import com.google.gson.Gson;


public class MifosRestClient implements RestClient {
    
    private final String baseURL;

    private final String userName;

    private final String password;

    private final String tenantId;

    private String authToken;
    
    public MifosRestClient() {
        baseURL = "http://localhost:8080/api/v1/"; // System.getProperty("mifos.endpoint");
        userName = "mifos"; // System.getProperty("mifos.user.id");
        password = "password"; // System.getProperty("mifos.password");
        tenantId = "default"; // System.getProperty("mifos.tenant.id");
    };

    public static final class Header {
        public static final String AUTHORIZATION = "Authorization";
        public static final String CONTENT_TYPE = "Content-Type";
        public static final String MIFOS_TENANT_ID = "X-Mifos-Platform-TenantId";
    }
    

    @Override
    public void post(String path, String payload) {
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

    private String createAuthToken() {
        String url = baseURL + "authentication?username=" + userName + "&password=" + password;
        try {

            SimpleHttpResponse response = new HttpRequestBuilder().withURL(url).withMethod(Method.POST)
                        .addHeader(Header.MIFOS_TENANT_ID, tenantId)
                        .addHeader(Header.CONTENT_TYPE, "application/json; charset=utf-8").execute();

            String content = readContentAndClose(response.getContent());
            AuthToken auth = new Gson().fromJson(content, AuthToken.class);
            return auth.getBase64EncodedAuthenticationKey();
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

}
