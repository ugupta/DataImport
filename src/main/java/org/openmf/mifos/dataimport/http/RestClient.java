package org.openmf.mifos.dataimport.http;


public interface RestClient {
    
    void post(String path, String payload);

}
