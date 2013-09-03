package org.openmf.mifos.dataimport.populator;

import java.io.IOException;

import org.openmf.mifos.dataimport.http.MifosRestClient;

public class WorkbookPopulatorFactory {
	
	
	  public static final WorkbookPopulator createWorkbookPopulator(String parameter, String template) throws IOException {
            
	        if(template.trim().equals("client")) 
	             return new ClientWorkbookPopulator (parameter, new MifosRestClient());
	        else if(template.trim().equals("loan"))
	        	 return new LoanWorkbookPopulator(new MifosRestClient());
	        else if(template.trim().equals("loanRepaymentHistory"))
	        	 return new LoanRepaymentWorkbookPopulator(new MifosRestClient());
	        else if(template.trim().equals("savings"))
	        	 return new SavingsWorkbookPopulator(new MifosRestClient());
	        throw new IllegalArgumentException(" Can't find populator. ");
	    }

}
