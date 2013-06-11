package org.openmf.mifos.dataimport.http;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;

import org.apache.commons.io.IOUtils;

public class SimpleHttpResponse
{

	private HttpURLConnection connection;
	
	private InputStream content;

	public SimpleHttpResponse(HttpURLConnection connection)
	{
		this.connection = connection;
	}

	public int getStatus() throws IOException
	{
		return connection.getResponseCode();
	}

	public InputStream getContent() throws IOException
	{
		content = connection.getInputStream();
		return content;
	}

	public void destroy()
	{
		IOUtils.closeQuietly(content);
	}

}
