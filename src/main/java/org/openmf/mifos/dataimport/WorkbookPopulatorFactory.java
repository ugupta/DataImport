package org.openmf.mifos.dataimport;

import java.io.IOException;

import org.openmf.mifos.dataimport.http.MifosRestClient;
import org.openmf.mifos.dataimport.web.DownloadServiceServlet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class WorkbookPopulatorFactory {
	
	private static final Logger logger = LoggerFactory.getLogger(WorkbookPopulatorFactory.class);
	
	  public static final WorkbookPopulator createWorkbookPopulator(String template) throws IOException {
            
	        if(template.trim().equals("client")) {
	             return new ClientWorkbookPopulator (new MifosRestClient());
	        }
	        throw new IllegalArgumentException(" Check ");
	    }

}
