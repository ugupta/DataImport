package org.openmf.mifos.dataimport.populator;

import java.io.IOException;

import org.openmf.mifos.dataimport.http.MifosRestClient;

public class WorkbookPopulatorFactory {
	
	
	  public static final WorkbookPopulator createWorkbookPopulator(String template) throws IOException {
            
	        if(template.trim().equals("client")) {
	             return new ClientWorkbookPopulator (new MifosRestClient());
	        }
	        throw new IllegalArgumentException(" Check ");
	    }

}
