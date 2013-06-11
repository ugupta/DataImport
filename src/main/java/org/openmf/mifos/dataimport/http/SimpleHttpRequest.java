package org.openmf.mifos.dataimport.http;

import java.io.IOException;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.io.IOUtils;

public class SimpleHttpRequest {

    private static final int HTTP_TIMEOUT = 100 * 1000; // 100 secs
	
    private final HttpURLConnection connection;

    public static enum Method {
        GET, POST;
    }
    
    public static final class Header {
    	public static final String AUTHORIZATION = "Authorization";
    	public static final String CONTENT_TYPE = "Content-Type";
    	public static final String MIFOS_TENANT_ID = "X-Mifos-Platform-TenantId";
    }

    public SimpleHttpRequest(String url, Method method, Map<String, String> headers, String content) throws IOException {
        URL obj = new URL(url);
        connection = (HttpURLConnection) obj.openConnection();
        connection.setRequestMethod(method.name());
        connection.setReadTimeout(HTTP_TIMEOUT);
        connection.setUseCaches(false);
        for (Entry<String, String> header : headers.entrySet()) {
            connection.addRequestProperty(header.getKey(), header.getValue());
        }
        if (content != null) {
            connection.setDoOutput(true);
            OutputStream out = connection.getOutputStream();
            IOUtils.write(content, out);
            out.flush();
            out.close();
        }
    }
    
    public SimpleHttpRequest(String url, Map<String, String> headers, String content) throws IOException {
        this(url, Method.GET, headers, content);
    }
    
    public SimpleHttpRequest(String url, String content) throws IOException {
        this(url, Method.GET, new HashMap<String, String>(), content);
    }
    
    public SimpleHttpRequest(String url) throws IOException {
        this(url, Method.GET, new HashMap<String, String>(), null);
    }

    public int status() throws IOException {
        return connection.getResponseCode();
    }

    public HttpURLConnection getConnection() throws IOException {
        return connection;
        
    }

}