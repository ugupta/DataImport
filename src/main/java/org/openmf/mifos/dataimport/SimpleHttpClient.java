package org.openmf.mifos.dataimport;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.io.IOUtils;

public class SimpleHttpClient {

    private final Method method;

    private final HttpURLConnection connection;

    private String content;

    private Map<String, String> headers = new HashMap<String, String>();

    public static enum Method {
        GET, POST;
    }

    public SimpleHttpClient(String url, Method method) throws IOException {
        this.method = method;
        URL obj = new URL(url);
        connection = (HttpURLConnection) obj.openConnection();
    }

    public SimpleHttpClient header(String name, String value) {
        headers.put(name, value);
        return this;
    }

    public SimpleHttpClient sendContent(String content) {
        this.content = content;
        return this;
    }

    public void execute() throws IOException {
        connection.setRequestMethod(method.name());
        connection.setReadTimeout(5000);
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

    public int status() throws IOException {
        return connection.getResponseCode();
    }

    public String getContent() throws IOException {
        InputStream is = connection.getInputStream();
        String content = IOUtils.toString(new InputStreamReader(is, "UTF-8"));
        is.close();
        return content;
        
    }

}
